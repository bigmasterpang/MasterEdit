//! 受限公式求值器（0.20.0「公式计算」）。
//!
//! 应用里**不写计算公式的结果**（只把公式文本写进文件，由 Excel / WPS 打开时算），
//! 但用户输入 `=SUM(E1:E2)` 后需要立刻看到结果。所以这里实现一个**受限但实用**的求值器：
//! 只覆盖日常够用的运算符、函数与引用，超出范围的一律给出**明确的中文错误**，
//! 由前端显示（并提示"用 Excel/WPS 打开可得到完整计算"），绝不 panic、绝不返回错误值。
//!
//! ## 支持范围
//! - 运算符：`+ - * / ^ %`（百分号）、一元 `-`/`+`、括号、比较 `= <> > < >= <=`、连接 `&`
//! - 函数：`SUM` `AVERAGE` `COUNT` `COUNTA` `MIN` `MAX` `ABS` `ROUND` `INT` `IF` `IFERROR`
//!   `AND` `OR` `NOT` `LEN` `LEFT` `RIGHT` `MID` `TRIM` `UPPER` `LOWER` `CONCAT`/`CONCATENATE`
//!   `TODAY` `NOW`
//! - 引用：`A1`、`$A$1`、区域 `A1:B2`、整列 `B:B`、整行 `2:2`、跨表 `Sheet1!A1`（表名可加
//!   单引号：`'我的 表'!A1`）
//!
//! ## 明确的边界（都会转成 `EvalResult.error`）
//! - 不支持的函数 → 「暂不支持函数 XXX」（不是 #NAME?，前端要原样展示这句）
//! - 除零 → `#DIV/0!`；类型不匹配 → `#VALUE!`；数值溢出/无意义 → `#NUM!`；
//!   表或单元格引用不存在 → `#REF!`
//! - 循环引用 → 「循环引用」；引用层级超过 [`MAX_DEPTH`] → 「公式引用层级过深」
//! - 区域出现在需要单值的位置（如 `=A1:B2+1`）→ `#VALUE!`（不做 Excel 的隐式交集）
//! - 区域超过 [`MAX_AREA_CELLS`] 格 → 「区域过大，暂不支持计算」
//!
//! ## 错误传播（与 Excel 一致）
//! 被引用区域里出现错误值时，**聚合函数把错误向上传播**，不会「跳过错误算出一个数字」：
//! - 单元格本身是错误（文件里缓存成 `#DIV/0!` 等，或它的公式**求值后**得到错误、
//!   或它的公式用了不支持的函数）→ `SUM`/`AVERAGE`/`COUNT`/`COUNTA`/`MIN`/`MAX`/`AND`/`OR`/
//!   `CONCAT` 的结果就是那个错误；
//! - 错误会带上**出错单元格的坐标**，方便用户定位：Excel 错误码写成
//!   `#DIV/0!（B5 出错）`，中文提示写成 `暂不支持函数 VLOOKUP（B5）`；
//! - 区域内多处错误时取**先行后列的第一个**；错误跨过单元格边界时**就近的根因优先**
//!   （A2 的公式因为 B1 出错而报错，传上去的仍是「B1 出错」）；
//! - 普通文本仍然按 Excel 语义忽略（`SUM` 跳过、`COUNT` 只数数字、`COUNTA` 数非空），
//!   空单元格同样忽略 —— 只有**错误**会传播；
//! - `IFERROR` 能照常捕获传播上来的错误（错误走的是普通错误通道，不是直接失败）。
//!
//! ## 取值口径（与网格显示保持一致）
//! - 数据来自 `office::read_workbook_bytes`（**影子优先**，企业加密文件在内存中解密），
//!   所以结构操作后的坐标是对的；
//! - `overrides`（前端尚未落盘的编辑）优先于文件内容；其中若是公式则**递归求值**；
//! - 文件里带公式的单元格**按公式重算**（我们写回时不算缓存值，直接读缓存会得到 0），
//!   因此 `=第二表!A2` 这种引用能拿到真实结果；
//! - 日期单元格按网格的显示口径当作**文本**（`2024-01-01`），`=A1+1` 会得到 `#VALUE!`
//!   —— 需要按日期算请用 Excel / WPS 打开；
//! - 聚合里的文本按「纯数字文本参与、其余忽略」处理（与 `office_write::parse_number` 一致）；
//!   布尔与空单元格不参与 SUM/AVERAGE/MIN/MAX/COUNT，但 COUNTA 会把它们算作非空。

use crate::commands::office;
use crate::commands::office_ops::{cell_display, format_float};
use crate::commands::office_write::{parse_bool, parse_number, CellEdit};
use calamine::{Data, Range, Reader, Sheets};
use std::cmp::Ordering;
use std::collections::HashMap;
use std::io::Cursor;
use std::path::Path;

/// 公式引用的最大嵌套层级（overrides 与文件内公式都算）
pub(crate) const MAX_DEPTH: usize = 10;
/// 单个区域允许展开的最大单元格数（整列/整行会先按已用区域收敛）
const MAX_AREA_CELLS: u64 = 1_000_000;

const ERR_VALUE: &str = "#VALUE!";
const ERR_DIV0: &str = "#DIV/0!";
const ERR_NUM: &str = "#NUM!";
const ERR_REF: &str = "#REF!";

/* ------------------------------------------------------------------ */
/* 命令的数据形状                                                      */
/* ------------------------------------------------------------------ */

/// 一批公式的求值请求（行/列是 0 起坐标）
#[derive(serde::Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct EvalRequest {
    pub row: u32,
    pub col: u32,
    pub formula: String,
}

/// 单条求值结果
#[derive(serde::Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct EvalResult {
    pub row: u32,
    pub col: u32,
    /// 计算成功的显示文本（数字不带多余小数、布尔为 TRUE/FALSE、文本原样）
    pub value: Option<String>,
    /// 计算失败的原因（中文，例如「暂不支持函数 VLOOKUP」），与 value 互斥
    pub error: Option<String>,
}

/* ------------------------------------------------------------------ */
/* 值                                                                  */
/* ------------------------------------------------------------------ */

#[derive(Clone, Debug, PartialEq)]
enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    Empty,
    /// 错误值。`at` 是**出错单元格**的 A1 坐标（1 起），None 表示还没定位 ——
    /// 错误跨过单元格边界（被区域/单元格引用读到）时会补上位置，见 [`Value::locate`]。
    Error { msg: String, at: Option<String> },
}

impl Value {
    /// 造一个错误值（尚无位置）
    fn error(msg: impl Into<String>) -> Value {
        Value::Error {
            msg: msg.into(),
            at: None,
        }
    }

    /// 给错误补上「它是在哪个单元格里出的」。
    ///
    /// **就近的根因优先**：已经有位置的错误保持原样（例如 A2 的公式是 `=SUM(B1:B3)`
    /// 且 B1 出错，那么一路传上去的仍是「B1 出错」，不会变成「A2 出错」而丢掉根因）。
    fn locate(self, at: String) -> Value {
        match self {
            Value::Error { msg, at: None } => Value::Error {
                msg,
                at: Some(at),
            },
            other => other,
        }
    }

    /// 显示文本（成功值用；错误值在 `EvalResult.error` 里）
    fn display(&self) -> String {
        match self {
            Value::Number(n) => format_float(*n),
            Value::Text(s) => s.clone(),
            Value::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            Value::Empty => String::new(),
            Value::Error { msg, .. } => msg.clone(),
        }
    }

    fn is_error(&self) -> bool {
        matches!(self, Value::Error { .. })
    }

    /// 错误文本（带位置）：Excel 错误码写成 `#DIV/0!（B5 出错）`，
    /// 中文提示写成 `暂不支持函数 VLOOKUP（B5）`
    fn error_text(&self) -> String {
        match self {
            Value::Error { msg, at } => match at {
                Some(at) if msg.starts_with('#') => format!("{msg}（{at} 出错）"),
                Some(at) => format!("{msg}（{at}）"),
                None => msg.clone(),
            },
            other => other.display(),
        }
    }
}

/// 0 起行列 → A1 坐标（1 起列字母 + 1 起行号），错误定位用
fn a1(row: u32, col: u32) -> String {
    let mut n = col + 1;
    let mut letters = String::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        letters.insert(0, (b'A' + rem as u8) as char);
        n = (n - 1) / 26;
    }
    format!("{letters}{}", row + 1)
}

/// calamine 单元格 → 求值用的值（口径与网格显示一致）
fn data_value(data: &Data) -> Value {
    match data {
        Data::Empty => Value::Empty,
        Data::String(s) => Value::Text(s.clone()),
        Data::Int(i) => Value::Number(*i as f64),
        Data::Float(f) => Value::Number(*f),
        Data::Bool(b) => Value::Bool(*b),
        // 日期/时长按网格里的显示文本参与（见模块注释的取舍）
        Data::DateTime(_) | Data::DateTimeIso(_) | Data::DurationIso(_) => {
            Value::Text(cell_display(data))
        }
        Data::Error(_) => Value::error(cell_display(data)),
    }
}

/// 当作数字用（Excel 的隐式转换：空=0、TRUE=1、纯数字文本=数字）
fn coerce_number(v: &Value) -> Result<f64, Value> {
    match v {
        Value::Number(n) => Ok(*n),
        Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Value::Empty => Ok(0.0),
        Value::Text(s) => parse_number(s).ok_or_else(|| Value::error(ERR_VALUE)),
        // 错误值原样往上传（带位置）
        Value::Error { .. } => Err(v.clone()),
    }
}

/// 当作文本用（数字按显示口径、布尔 TRUE/FALSE）
fn as_text(v: &Value) -> Result<String, Value> {
    match v {
        Value::Error { .. } => Err(v.clone()),
        other => Ok(other.display()),
    }
}

/// 当作布尔用（用于 IF / AND / OR / NOT 的条件）
fn as_bool(v: &Value) -> Result<bool, Value> {
    match v {
        Value::Bool(b) => Ok(*b),
        Value::Number(n) => Ok(*n != 0.0),
        Value::Empty => Ok(false),
        Value::Text(s) => parse_number(s)
            .map(|n| n != 0.0)
            .ok_or_else(|| Value::error(ERR_VALUE)),
        Value::Error { .. } => Err(v.clone()),
    }
}

/* ------------------------------------------------------------------ */
/* 词法分析                                                            */
/* ------------------------------------------------------------------ */

#[derive(Clone, Debug, PartialEq)]
enum Token {
    Number(f64),
    /// 双引号字符串字面量
    Text(String),
    /// 标识符：函数名 / 单元格引用 / 工作表名
    Ident(String),
    Op(String),
    LParen,
    RParen,
    Comma,
    Colon,
    Bang,
}

fn is_ident_start(c: char) -> bool {
    c.is_alphabetic() || c == '_' || c == '$'
}

fn is_ident_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '$' || c == '.'
}

fn tokenize(src: &str) -> Result<Vec<Token>, String> {
    let chars: Vec<char> = src.chars().collect();
    let mut i = 0usize;
    let mut out = Vec::new();
    while i < chars.len() {
        let c = chars[i];
        if c.is_whitespace() {
            i += 1;
            continue;
        }
        match c {
            '(' => {
                out.push(Token::LParen);
                i += 1;
            }
            ')' => {
                out.push(Token::RParen);
                i += 1;
            }
            ',' => {
                out.push(Token::Comma);
                i += 1;
            }
            ':' => {
                out.push(Token::Colon);
                i += 1;
            }
            '!' => {
                out.push(Token::Bang);
                i += 1;
            }
            '"' => {
                // 字符串字面量：内部 "" 表示一个引号
                let mut s = String::new();
                i += 1;
                loop {
                    if i >= chars.len() {
                        return Err("字符串缺少右引号".to_string());
                    }
                    if chars[i] == '"' {
                        if i + 1 < chars.len() && chars[i + 1] == '"' {
                            s.push('"');
                            i += 2;
                            continue;
                        }
                        i += 1;
                        break;
                    }
                    s.push(chars[i]);
                    i += 1;
                }
                out.push(Token::Text(s));
            }
            '\'' => {
                // 带引号的工作表名：'我的 表'
                let mut s = String::new();
                i += 1;
                loop {
                    if i >= chars.len() {
                        return Err("工作表名缺少右引号".to_string());
                    }
                    if chars[i] == '\'' {
                        if i + 1 < chars.len() && chars[i + 1] == '\'' {
                            s.push('\'');
                            i += 2;
                            continue;
                        }
                        i += 1;
                        break;
                    }
                    s.push(chars[i]);
                    i += 1;
                }
                out.push(Token::Ident(s));
            }
            _ if c.is_ascii_digit() || c == '.' => {
                let mut s = String::new();
                while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                    s.push(chars[i]);
                    i += 1;
                }
                let n: f64 = s
                    .parse()
                    .map_err(|_| format!("无法解析数字「{s}」"))?;
                out.push(Token::Number(n));
            }
            _ if is_ident_start(c) => {
                let mut s = String::new();
                while i < chars.len() && is_ident_char(chars[i]) {
                    s.push(chars[i]);
                    i += 1;
                }
                out.push(Token::Ident(s));
            }
            '+' | '-' | '*' | '/' | '^' | '%' | '&' | '=' => {
                out.push(Token::Op(c.to_string()));
                i += 1;
            }
            '<' => {
                if i + 1 < chars.len() && chars[i + 1] == '>' {
                    out.push(Token::Op("<>".to_string()));
                    i += 2;
                } else if i + 1 < chars.len() && chars[i + 1] == '=' {
                    out.push(Token::Op("<=".to_string()));
                    i += 2;
                } else {
                    out.push(Token::Op("<".to_string()));
                    i += 1;
                }
            }
            '>' => {
                if i + 1 < chars.len() && chars[i + 1] == '=' {
                    out.push(Token::Op(">=".to_string()));
                    i += 2;
                } else {
                    out.push(Token::Op(">".to_string()));
                    i += 1;
                }
            }
            other => return Err(format!("无法识别的字符「{other}」")),
        }
    }
    Ok(out)
}

/* ------------------------------------------------------------------ */
/* 引用解析                                                            */
/* ------------------------------------------------------------------ */

/// 一个矩形区域（0 起）。`row1`/`col1` 为 None 表示「到已用区域末尾」：
/// 整列 `B:B`、整行 `2:2`、多列 `B:C` 都靠它表达。
#[derive(Clone, Debug, PartialEq)]
struct Area {
    row0: u32,
    col0: u32,
    row1: Option<u32>,
    col1: Option<u32>,
}

impl Area {
    fn cell(row: u32, col: u32) -> Area {
        Area {
            row0: row,
            col0: col,
            row1: Some(row),
            col1: Some(col),
        }
    }

    fn is_single(&self) -> bool {
        self.row1 == Some(self.row0) && self.col1 == Some(self.col0)
    }
}

/// 列字母 → 0 起的列号（A=0、Z=25、AA=26）
fn column_index(letters: &str) -> Option<u32> {
    let mut n: u32 = 0;
    for c in letters.to_ascii_uppercase().chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        let d = (c as u8 - b'A') as u32 + 1;
        n = n.checked_mul(26)?.checked_add(d)?;
    }
    if n == 0 {
        None
    } else {
        Some(n - 1)
    }
}

/// `A1` / `$A$1` → (列, 行)（0 起）；不是单元格引用返回 None
fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
    let cleaned: String = s.chars().filter(|c| *c != '$').collect();
    if cleaned.is_empty() {
        return None;
    }
    let split = cleaned
        .char_indices()
        .find(|(_, c)| c.is_ascii_digit())
        .map(|(i, _)| i)?;
    let (letters, digits) = cleaned.split_at(split);
    if !letters.chars().all(|c| c.is_ascii_alphabetic())
        || !digits.chars().all(|c| c.is_ascii_digit())
    {
        return None;
    }
    let col = column_index(letters)?;
    let row: u32 = digits.parse().ok()?;
    if row == 0 {
        return None;
    }
    Some((col, row - 1))
}

/// `B` / `$B` → 0 起的列号（整列引用用）
fn parse_col_ref(s: &str) -> Option<u32> {
    let cleaned: String = s.chars().filter(|c| *c != '$').collect();
    if cleaned.is_empty() || !cleaned.chars().all(|c| c.is_ascii_alphabetic()) {
        return None;
    }
    column_index(&cleaned)
}

/* ------------------------------------------------------------------ */
/* 语法分析                                                            */
/* ------------------------------------------------------------------ */

#[derive(Clone, Copy, Debug, PartialEq)]
enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Eq,
    Ne,
    Gt,
    Lt,
    Ge,
    Le,
}

#[derive(Clone, Debug, PartialEq)]
enum Expr {
    Number(f64),
    Text(String),
    Bool(bool),
    Ref {
        sheet: Option<String>,
        area: Area,
    },
    Unary(char, Box<Expr>),
    Percent(Box<Expr>),
    Binary(BinOp, Box<Expr>, Box<Expr>),
    Call(String, Vec<Expr>),
}

struct Parser<'a> {
    tokens: &'a [Token],
    pos: usize,
}

impl<'a> Parser<'a> {
    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.pos)
    }

    fn next(&mut self) -> Option<Token> {
        let t = self.tokens.get(self.pos).cloned();
        if t.is_some() {
            self.pos += 1;
        }
        t
    }

    fn peek_op(&self) -> Option<&str> {
        match self.peek() {
            Some(Token::Op(op)) => Some(op.as_str()),
            _ => None,
        }
    }

    fn parse_all(mut self) -> Result<Expr, String> {
        let expr = self.comparison()?;
        if self.pos < self.tokens.len() {
            return Err("公式语法错误：表达式末尾有多余内容".to_string());
        }
        Ok(expr)
    }

    fn comparison(&mut self) -> Result<Expr, String> {
        let mut left = self.concat()?;
        loop {
            let op = match self.peek_op() {
                Some("=") => BinOp::Eq,
                Some("<>") => BinOp::Ne,
                Some(">") => BinOp::Gt,
                Some("<") => BinOp::Lt,
                Some(">=") => BinOp::Ge,
                Some("<=") => BinOp::Le,
                _ => break,
            };
            self.pos += 1;
            let right = self.concat()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn concat(&mut self) -> Result<Expr, String> {
        let mut left = self.additive()?;
        while self.peek_op() == Some("&") {
            self.pos += 1;
            let right = self.additive()?;
            left = Expr::Binary(BinOp::Concat, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn additive(&mut self) -> Result<Expr, String> {
        let mut left = self.multiplicative()?;
        loop {
            let op = match self.peek_op() {
                Some("+") => BinOp::Add,
                Some("-") => BinOp::Sub,
                _ => break,
            };
            self.pos += 1;
            let right = self.multiplicative()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn multiplicative(&mut self) -> Result<Expr, String> {
        let mut left = self.unary()?;
        loop {
            let op = match self.peek_op() {
                Some("*") => BinOp::Mul,
                Some("/") => BinOp::Div,
                _ => break,
            };
            self.pos += 1;
            let right = self.unary()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn unary(&mut self) -> Result<Expr, String> {
        match self.peek_op() {
            Some("-") => {
                self.pos += 1;
                Ok(Expr::Unary('-', Box::new(self.unary()?)))
            }
            Some("+") => {
                self.pos += 1;
                self.unary()
            }
            _ => self.power(),
        }
    }

    fn power(&mut self) -> Result<Expr, String> {
        let base = self.postfix()?;
        if self.peek_op() == Some("^") {
            self.pos += 1;
            // 右结合：2^3^2 = 2^(3^2)
            let exp = self.unary()?;
            return Ok(Expr::Binary(BinOp::Pow, Box::new(base), Box::new(exp)));
        }
        Ok(base)
    }

    fn postfix(&mut self) -> Result<Expr, String> {
        let mut expr = self.primary()?;
        while self.peek_op() == Some("%") {
            self.pos += 1;
            expr = Expr::Percent(Box::new(expr));
        }
        Ok(expr)
    }

    fn primary(&mut self) -> Result<Expr, String> {
        match self.next() {
            Some(Token::Number(n)) => {
                // 整行区域：2:2 / 2:3（第一个数字是行号）
                if matches!(self.peek(), Some(Token::Colon)) {
                    if let Some(Token::Number(m)) = self.tokens.get(self.pos + 1).cloned() {
                        if n.fract() == 0.0 && m.fract() == 0.0 && n >= 1.0 && m >= 1.0 {
                            self.pos += 2;
                            return Ok(Expr::Ref {
                                sheet: None,
                                area: Area {
                                    row0: (n as u32) - 1,
                                    col0: 0,
                                    row1: Some((m as u32) - 1),
                                    col1: None,
                                },
                            });
                        }
                    }
                }
                Ok(Expr::Number(n))
            }
            Some(Token::Text(s)) => Ok(Expr::Text(s)),
            Some(Token::LParen) => {
                let expr = self.comparison()?;
                match self.next() {
                    Some(Token::RParen) => Ok(expr),
                    _ => Err("公式语法错误：缺少右括号".to_string()),
                }
            }
            Some(Token::Ident(name)) => self.primary_from_ident(name),
            _ => Err("公式语法错误：表达式不完整".to_string()),
        }
    }

    /// 标识符后面可能跟：`(` 函数、`!` 跨表引用、`:` 区域
    fn primary_from_ident(&mut self, name: String) -> Result<Expr, String> {
        // TRUE / FALSE
        match name.to_ascii_uppercase().as_str() {
            "TRUE" => return Ok(Expr::Bool(true)),
            "FALSE" => return Ok(Expr::Bool(false)),
            _ => {}
        }
        // 函数调用
        if matches!(self.peek(), Some(Token::LParen)) {
            self.pos += 1;
            let mut args = Vec::new();
            if matches!(self.peek(), Some(Token::RParen)) {
                self.pos += 1;
                return Ok(Expr::Call(name.to_ascii_uppercase(), args));
            }
            loop {
                args.push(self.comparison()?);
                match self.next() {
                    Some(Token::Comma) => continue,
                    Some(Token::RParen) => break,
                    _ => return Err("公式语法错误：函数参数缺少右括号或逗号".to_string()),
                }
            }
            return Ok(Expr::Call(name.to_ascii_uppercase(), args));
        }
        // 跨表引用：Sheet1!A1 / '我的 表'!A1
        let mut sheet = None;
        let mut token = Token::Ident(name);
        if matches!(self.peek(), Some(Token::Bang)) {
            self.pos += 1;
            match &token {
                Token::Ident(s) => sheet = Some(s.clone()),
                _ => return Err("公式语法错误：工作表名无效".to_string()),
            }
            token = match self.next() {
                Some(t) => t,
                None => return Err("公式语法错误：跨表引用不完整".to_string()),
            };
        }
        let end = self.parse_ref_tail(&token, sheet)?;
        Ok(end)
    }

    /// `A1` / `B` / `2:2` / `A1:B2` / `B:C`（`sheet!` 前缀已解析）
    fn parse_ref_tail(&mut self, token: &Token, sheet: Option<String>) -> Result<Expr, String> {
        let mut area = single_token_area(token, sheet.is_some())?;
        if matches!(self.peek(), Some(Token::Colon)) {
            self.pos += 1;
            let second = match self.next() {
                Some(t) => t,
                None => return Err("公式语法错误：区域不完整".to_string()),
            };
            let tail = single_token_area(&second, sheet.is_some())?;
            area = Area {
                row0: area.row0.min(tail.row0),
                col0: area.col0.min(tail.col0),
                row1: match (area.row1, tail.row1) {
                    (Some(a), Some(b)) => Some(a.max(b)),
                    _ => None,
                },
                col1: match (area.col1, tail.col1) {
                    (Some(a), Some(b)) => Some(a.max(b)),
                    _ => None,
                },
            };
        }
        Ok(Expr::Ref { sheet, area })
    }
}

/// 把一个 token 解析成「区域片段」：单元格 / 整列 / 整行
fn single_token_area(token: &Token, _qualified: bool) -> Result<Area, String> {
    match token {
        Token::Ident(name) => {
            if let Some((col, row)) = parse_cell_ref(name) {
                return Ok(Area::cell(row, col));
            }
            if let Some(col) = parse_col_ref(name) {
                // 整列：B:B / B（在区域里）
                return Ok(Area {
                    row0: 0,
                    col0: col,
                    row1: None,
                    col1: Some(col),
                });
            }
            Err(format!("#NAME?（无法识别引用「{name}」）"))
        }
        Token::Number(n) => {
            if n.fract() == 0.0 && *n >= 1.0 {
                Ok(Area {
                    row0: (*n as u32) - 1,
                    col0: 0,
                    row1: Some((*n as u32) - 1),
                    col1: None,
                })
            } else {
                Err("公式语法错误：区域坐标无效".to_string())
            }
        }
        _ => Err("公式语法错误：区域坐标无效".to_string()),
    }
}

fn parse_formula(formula: &str) -> Result<Expr, String> {
    let text = formula.trim();
    let text = text.strip_prefix('=').unwrap_or(text).trim();
    if text.is_empty() {
        return Err("公式为空".to_string());
    }
    let tokens = tokenize(text)?;
    if tokens.is_empty() {
        return Err("公式为空".to_string());
    }
    Parser { tokens: &tokens, pos: 0 }.parse_all()
}

/* ------------------------------------------------------------------ */
/* 求值                                                                */
/* ------------------------------------------------------------------ */

#[derive(Clone)]
enum Override {
    Value(Value),
    Formula(String),
}

struct Book<'a> {
    sheets: Sheets<Cursor<&'a [u8]>>,
    values: HashMap<String, Option<Range<Data>>>,
    formulas: HashMap<String, Option<Range<String>>>,
    overrides: HashMap<(String, u32, u32), Override>,
    depth: usize,
    visiting: Vec<(String, u32, u32)>,
}

impl<'a> Book<'a> {
    /// 按需解析某个工作表（值 + 公式），并缓存
    fn ensure(&mut self, sheet: &str) {
        let key = sheet.to_lowercase();
        if self.values.contains_key(&key) {
            return;
        }
        let values = self.sheets.worksheet_range(sheet).ok();
        let formulas = self.sheets.worksheet_formula(sheet).ok();
        self.values.insert(key.clone(), values);
        self.formulas.insert(key, formulas);
    }

    fn has_sheet(&mut self, sheet: &str) -> bool {
        self.ensure(sheet);
        self.values
            .get(&sheet.to_lowercase())
            .map(|v| v.is_some())
            .unwrap_or(false)
    }

    /// 已用区域的右下角（0 起）；整列/整行引用按它收敛
    fn used_end(&mut self, sheet: &str) -> (u32, u32) {
        self.ensure(sheet);
        self.values
            .get(&sheet.to_lowercase())
            .and_then(|v| v.as_ref())
            .and_then(|r| r.end())
            .unwrap_or((0, 0))
    }

    fn raw(&mut self, sheet: &str, row: u32, col: u32) -> Value {
        self.ensure(sheet);
        let data = self
            .values
            .get(&sheet.to_lowercase())
            .and_then(|v| v.as_ref())
            .and_then(|r| r.get_value((row, col)))
            .cloned();
        match data {
            Some(d) => data_value(&d),
            None => Value::Empty,
        }
    }

    fn formula_at(&mut self, sheet: &str, row: u32, col: u32) -> Option<String> {
        self.ensure(sheet);
        self.formulas
            .get(&sheet.to_lowercase())
            .and_then(|v| v.as_ref())
            .and_then(|r| r.get_value((row, col)))
            .map(|s| s.trim().trim_start_matches('=').trim().to_string())
            .filter(|s| !s.is_empty())
    }

    /// 取一个单元格的值：overrides（可能递归）→ 文件里的公式（重算）→ 缓存值
    fn cell(&mut self, sheet: &str, row: u32, col: u32) -> Value {
        if let Some(ov) = self
            .overrides
            .get(&(sheet.to_lowercase(), row, col))
            .cloned()
        {
            return match ov {
                Override::Value(v) => v,
                Override::Formula(f) => self.eval_at(sheet, row, col, &f),
            };
        }
        if let Some(f) = self.formula_at(sheet, row, col) {
            return self.eval_at(sheet, row, col, &f);
        }
        self.raw(sheet, row, col)
    }

    /// 在「某单元格正在被求值」的上下文里求值一个公式（带循环引用与深度保护）
    fn eval_at(&mut self, sheet: &str, row: u32, col: u32, formula: &str) -> Value {
        let key = (sheet.to_lowercase(), row, col);
        if self.visiting.iter().any(|k| *k == key) {
            return Value::error("循环引用");
        }
        if self.depth >= MAX_DEPTH {
            return Value::error("公式引用层级过深");
        }
        let expr = match parse_formula(formula) {
            Ok(expr) => expr,
            Err(msg) => return Value::error(msg),
        };
        self.visiting.push(key);
        self.depth += 1;
        let out = self.eval(sheet, &expr);
        self.depth -= 1;
        self.visiting.pop();
        out
    }

    fn eval(&mut self, sheet: &str, expr: &Expr) -> Value {
        match expr {
            Expr::Number(n) => Value::Number(*n),
            Expr::Text(s) => Value::Text(s.clone()),
            Expr::Bool(b) => Value::Bool(*b),
            Expr::Ref { sheet: s, area } => {
                let target = s.clone().unwrap_or_else(|| sheet.to_string());
                if !self.has_sheet(&target) {
                    return Value::error(ERR_REF);
                }
                if area.is_single() {
                    // 单格引用：错误值补上该格的坐标，用户才知道是哪一格出的错
                    self.cell(&target, area.row0, area.col0)
                        .locate(a1(area.row0, area.col0))
                } else {
                    // 区域出现在需要单值的位置：不做 Excel 的隐式交集
                    Value::error(ERR_VALUE)
                }
            }
            Expr::Unary(op, inner) => {
                let v = self.eval(sheet, inner);
                if v.is_error() {
                    return v;
                }
                match coerce_number(&v) {
                    Ok(n) => {
                        let out = if *op == '-' { -n } else { n };
                        if out.is_finite() {
                            Value::Number(out)
                        } else {
                            Value::error(ERR_NUM)
                        }
                    }
                    Err(e) => e,
                }
            }
            Expr::Percent(inner) => {
                let v = self.eval(sheet, inner);
                if v.is_error() {
                    return v;
                }
                match coerce_number(&v) {
                    Ok(n) => Value::Number(n / 100.0),
                    Err(e) => e,
                }
            }
            Expr::Binary(op, left, right) => self.binary(sheet, *op, left, right),
            Expr::Call(name, args) => self.call(sheet, name, args),
        }
    }

    fn binary(&mut self, sheet: &str, op: BinOp, left: &Expr, right: &Expr) -> Value {
        let a = self.eval(sheet, left);
        if a.is_error() {
            return a;
        }
        let b = self.eval(sheet, right);
        if b.is_error() {
            return b;
        }
        match op {
            BinOp::Concat => Value::Text(format!("{}{}", a.display(), b.display())),
            BinOp::Eq | BinOp::Ne | BinOp::Gt | BinOp::Lt | BinOp::Ge | BinOp::Le => {
                Value::Bool(compare(op, &a, &b))
            }
            _ => {
                let x = match coerce_number(&a) {
                    Ok(v) => v,
                    Err(e) => return e,
                };
                let y = match coerce_number(&b) {
                    Ok(v) => v,
                    Err(e) => return e,
                };
                let out = match op {
                    BinOp::Add => x + y,
                    BinOp::Sub => x - y,
                    BinOp::Mul => x * y,
                    BinOp::Div => {
                        if y == 0.0 {
                            return Value::error(ERR_DIV0);
                        }
                        x / y
                    }
                    BinOp::Pow => x.powf(y),
                    _ => unreachable!("比较与连接已在上面处理"),
                };
                if out.is_finite() {
                    Value::Number(out)
                } else {
                    Value::error(ERR_NUM)
                }
            }
        }
    }

    /// 区域/参数展开成一串值（聚合函数用）。
    ///
    /// 区域内出现**错误值**时直接返回那个错误（带出错单元格坐标），也就是把错误向上传播 ——
    /// 与 Excel 一致：`SUM`/`AVERAGE`/`COUNT`/`MIN`/`MAX`/`COUNTA` 等都不会「跳过」错误算出
    /// 一个数字。普通文本仍然按 Excel 语义忽略（见 [`Book::aggregate`]），空单元格同样忽略。
    fn collect(&mut self, sheet: &str, args: &[Expr]) -> Result<Vec<Value>, Value> {
        let mut out = Vec::new();
        for arg in args {
            match arg {
                Expr::Ref { sheet: s, area } => {
                    let target = s.clone().unwrap_or_else(|| sheet.to_string());
                    out.extend(self.area_values(&target, area)?);
                }
                _ => {
                    let v = self.eval(sheet, arg);
                    if v.is_error() {
                        return Err(v);
                    }
                    out.push(v);
                }
            }
        }
        Ok(out)
    }

    /// 展开区域；遇到错误值立即返回（先行后列，所以就是区域里出现的第一个错误）
    fn area_values(&mut self, sheet: &str, area: &Area) -> Result<Vec<Value>, Value> {
        if !self.has_sheet(sheet) {
            return Err(Value::error(ERR_REF));
        }
        let (used_row, used_col) = self.used_end(sheet);
        let row1 = area.row1.unwrap_or(used_row);
        let col1 = area.col1.unwrap_or(used_col);
        let (r0, r1) = if area.row0 <= row1 {
            (area.row0, row1)
        } else {
            (row1, area.row0)
        };
        let (c0, c1) = if area.col0 <= col1 {
            (area.col0, col1)
        } else {
            (col1, area.col0)
        };
        let cells = (r1 - r0 + 1) as u64 * (c1 - c0 + 1) as u64;
        if cells > MAX_AREA_CELLS {
            return Err(Value::error("区域过大，暂不支持计算"));
        }
        let mut out = Vec::with_capacity(cells as usize);
        for r in r0..=r1 {
            for c in c0..=c1 {
                // cell() 会递归求值该格的公式（深度上限与循环引用检测都在里面），
                // 所以「公式用到了不支持的函数」「公式算出除零」都会在这里变成错误值
                let value = self.cell(sheet, r, c);
                if value.is_error() {
                    return Err(value.locate(a1(r, c)));
                }
                out.push(value);
            }
        }
        Ok(out)
    }

    fn call(&mut self, sheet: &str, name: &str, args: &[Expr]) -> Value {
        match name {
            "SUM" | "AVERAGE" | "COUNT" | "COUNTA" | "MIN" | "MAX" => {
                self.aggregate(sheet, name, args)
            }
            "IF" => {
                if args.len() < 2 || args.len() > 3 {
                    return Value::error("IF 需要 2 或 3 个参数");
                }
                let cond = self.eval(sheet, &args[0]);
                let truth = match as_bool(&cond) {
                    Ok(b) => b,
                    Err(e) => return e,
                };
                if truth {
                    self.eval(sheet, &args[1])
                } else if args.len() == 3 {
                    self.eval(sheet, &args[2])
                } else {
                    Value::Bool(false)
                }
            }
            "IFERROR" => {
                if args.len() != 2 {
                    return Value::error("IFERROR 需要 2 个参数");
                }
                let v = self.eval(sheet, &args[0]);
                match v {
                    // 被捕获的错误可能是「区域里某格出错」传播上来的（带位置），
                    // 这里照样接住 —— 错误走的是普通错误通道，不是直接失败
                    Value::Error { .. } => self.eval(sheet, &args[1]),
                    other => other,
                }
            }
            "AND" | "OR" => {
                let items = match self.collect(sheet, args) {
                    Ok(v) => v,
                    Err(e) => return e,
                };
                let mut seen = false;
                let mut all = true;
                let mut any = false;
                for item in &items {
                    if matches!(item, Value::Empty) {
                        continue;
                    }
                    let b = match as_bool(item) {
                        Ok(b) => b,
                        Err(e) => return e,
                    };
                    seen = true;
                    all &= b;
                    any |= b;
                }
                if !seen {
                    return Value::error(ERR_VALUE);
                }
                Value::Bool(if name == "AND" { all } else { any })
            }
            "NOT" => match self.scalar_arg(sheet, name, args) {
                Ok(v) => match as_bool(&v) {
                    Ok(b) => Value::Bool(!b),
                    Err(e) => e,
                },
                Err(e) => e,
            },
            "ABS" | "INT" => match self.number_arg(sheet, name, args) {
                Ok(n) => {
                    let out = if name == "ABS" { n.abs() } else { n.floor() };
                    if out.is_finite() {
                        Value::Number(out)
                    } else {
                        Value::error(ERR_NUM)
                    }
                }
                Err(e) => e,
            },
            "ROUND" => {
                if args.is_empty() || args.len() > 2 {
                    return Value::error("ROUND 需要 1 或 2 个参数");
                }
                let n = match self.number_arg(sheet, "ROUND", &args[..1]) {
                    Ok(n) => n,
                    Err(e) => return e,
                };
                let digits = if args.len() == 2 {
                    match self.number_arg(sheet, "ROUND", &args[1..2]) {
                        Ok(d) => d.trunc(),
                        Err(e) => return e,
                    }
                } else {
                    0.0
                };
                let factor = 10f64.powf(digits);
                let out = (n * factor).round() / factor;
                if out.is_finite() {
                    Value::Number(out)
                } else {
                    Value::error(ERR_NUM)
                }
            }
            "LEN" => match self.text_arg(sheet, name, args) {
                Ok(s) => Value::Number(s.chars().count() as f64),
                Err(e) => e,
            },
            "TRIM" => match self.text_arg(sheet, name, args) {
                // Excel 的 TRIM：去掉首尾空格，并把中间连续空格压成一个
                Ok(s) => Value::Text(s.split_whitespace().collect::<Vec<_>>().join(" ")),
                Err(e) => e,
            },
            "UPPER" => match self.text_arg(sheet, name, args) {
                Ok(s) => Value::Text(s.to_uppercase()),
                Err(e) => e,
            },
            "LOWER" => match self.text_arg(sheet, name, args) {
                Ok(s) => Value::Text(s.to_lowercase()),
                Err(e) => e,
            },
            "LEFT" | "RIGHT" => {
                if args.is_empty() || args.len() > 2 {
                    return Value::error(format!("{name} 需要 1 或 2 个参数"));
                }
                let text = match self.text_arg(sheet, name, &args[..1]) {
                    Ok(s) => s,
                    Err(e) => return e,
                };
                let count = if args.len() == 2 {
                    match self.number_arg(sheet, name, &args[1..2]) {
                        Ok(n) => n.trunc(),
                        Err(e) => return e,
                    }
                } else {
                    1.0
                };
                if count < 0.0 {
                    return Value::error(ERR_VALUE);
                }
                let n = count as usize;
                let chars: Vec<char> = text.chars().collect();
                let taken: String = if name == "LEFT" {
                    chars.iter().take(n).collect()
                } else {
                    chars.iter().skip(chars.len().saturating_sub(n)).collect()
                };
                Value::Text(taken)
            }
            "MID" => {
                if args.len() != 3 {
                    return Value::error("MID 需要 3 个参数");
                }
                let text = match self.text_arg(sheet, name, &args[..1]) {
                    Ok(s) => s,
                    Err(e) => return e,
                };
                let start = match self.number_arg(sheet, name, &args[1..2]) {
                    Ok(n) => n.trunc(),
                    Err(e) => return e,
                };
                let len = match self.number_arg(sheet, name, &args[2..3]) {
                    Ok(n) => n.trunc(),
                    Err(e) => return e,
                };
                if start < 1.0 || len < 0.0 {
                    return Value::error(ERR_VALUE);
                }
                let chars: Vec<char> = text.chars().collect();
                let from = (start as usize) - 1;
                let taken: String = chars.iter().skip(from).take(len as usize).collect();
                Value::Text(taken)
            }
            "CONCAT" | "CONCATENATE" => {
                let items = match self.collect(sheet, args) {
                    Ok(v) => v,
                    Err(e) => return e,
                };
                let mut out = String::new();
                for item in &items {
                    match as_text(item) {
                        Ok(s) => out.push_str(&s),
                        Err(e) => return e,
                    }
                }
                Value::Text(out)
            }
            "TODAY" => {
                let (y, m, d, _, _, _) = local_now_parts();
                Value::Text(format!("{y:04}-{m:02}-{d:02}"))
            }
            "NOW" => {
                let (y, m, d, hh, mm, ss) = local_now_parts();
                Value::Text(format!("{y:04}-{m:02}-{d:02} {hh:02}:{mm:02}:{ss:02}"))
            }
            other => Value::error(format!("暂不支持函数 {other}")),
        }
    }

    /// 单值参数（要求恰好 1 个，且不能是区域）
    fn scalar_arg(&mut self, sheet: &str, name: &str, args: &[Expr]) -> Result<Value, Value> {
        if args.len() != 1 {
            return Err(Value::error(format!("{name} 需要 1 个参数")));
        }
        let v = self.eval(sheet, &args[0]);
        if v.is_error() {
            return Err(v);
        }
        Ok(v)
    }

    fn number_arg(&mut self, sheet: &str, name: &str, args: &[Expr]) -> Result<f64, Value> {
        let v = self.scalar_arg(sheet, name, args)?;
        coerce_number(&v)
    }

    fn text_arg(&mut self, sheet: &str, name: &str, args: &[Expr]) -> Result<String, Value> {
        let v = self.scalar_arg(sheet, name, args)?;
        as_text(&v)
    }

    fn aggregate(&mut self, sheet: &str, name: &str, args: &[Expr]) -> Value {
        // collect 会在遇到错误值时直接返回它（带出错单元格坐标）—— 错误向上传播
        let items = match self.collect(sheet, args) {
            Ok(v) => v,
            Err(e) => return e,
        };
        let non_empty = items.iter().filter(|v| !matches!(v, Value::Empty)).count();
        // 数字：Number 本身 + 纯数字文本；布尔/空/普通文本都不参与（见模块注释）
        let numbers: Vec<f64> = items
            .iter()
            .filter_map(|v| match v {
                Value::Number(n) => Some(*n),
                Value::Text(s) => parse_number(s),
                _ => None,
            })
            .collect();
        match name {
            "SUM" => Value::Number(numbers.iter().sum()),
            "COUNT" => Value::Number(numbers.len() as f64),
            "COUNTA" => Value::Number(non_empty as f64),
            "AVERAGE" => {
                if numbers.is_empty() {
                    Value::error(ERR_DIV0)
                } else {
                    Value::Number(numbers.iter().sum::<f64>() / numbers.len() as f64)
                }
            }
            "MIN" => Value::Number(
                numbers
                    .iter()
                    .copied()
                    .fold(None::<f64>, |acc, n| Some(acc.map_or(n, |a: f64| a.min(n))))
                    .unwrap_or(0.0),
            ),
            "MAX" => Value::Number(
                numbers
                    .iter()
                    .copied()
                    .fold(None::<f64>, |acc, n| Some(acc.map_or(n, |a: f64| a.max(n))))
                    .unwrap_or(0.0),
            ),
            _ => Value::error(ERR_VALUE),
        }
    }
}

/// 比较运算：两边都能当数字就按数字比，否则按显示文本（不区分大小写）比
fn compare(op: BinOp, a: &Value, b: &Value) -> bool {
    let ord = match (coerce_number(a), coerce_number(b)) {
        (Ok(x), Ok(y)) => x.partial_cmp(&y).unwrap_or(Ordering::Equal),
        _ => a
            .display()
            .to_lowercase()
            .cmp(&b.display().to_lowercase()),
    };
    match op {
        BinOp::Eq => ord == Ordering::Equal,
        BinOp::Ne => ord != Ordering::Equal,
        BinOp::Gt => ord == Ordering::Greater,
        BinOp::Lt => ord == Ordering::Less,
        BinOp::Ge => ord != Ordering::Less,
        BinOp::Le => ord != Ordering::Greater,
        _ => false,
    }
}

/* ------------------------------------------------------------------ */
/* 当前本地时间（TODAY / NOW）                                          */
/* ------------------------------------------------------------------ */

/// 本地时间的 (年, 月, 日, 时, 分, 秒)。
///
/// Windows 上用 `GetLocalTime`（用户看到的日期必须是本地的，UTC 在 UTC+8 会让
/// 凌晨 0-8 点的 `TODAY()` 差一天）；其它平台退回 UTC —— 本项目只发布 Windows 版。
#[cfg(windows)]
fn local_now_parts() -> (i32, u32, u32, u32, u32, u32) {
    use windows::Win32::System::SystemInformation::GetLocalTime;
    let t = unsafe { GetLocalTime() };
    (
        t.wYear as i32,
        t.wMonth as u32,
        t.wDay as u32,
        t.wHour as u32,
        t.wMinute as u32,
        t.wSecond as u32,
    )
}

#[cfg(not(windows))]
fn local_now_parts() -> (i32, u32, u32, u32, u32, u32) {
    let secs = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0);
    let days = secs.div_euclid(86_400);
    let rem = secs.rem_euclid(86_400);
    let (y, m, d) = crate::commands::image::civil_from_days(days);
    (
        y as i32,
        m,
        d,
        (rem / 3600) as u32,
        ((rem % 3600) / 60) as u32,
        (rem % 60) as u32,
    )
}

/* ------------------------------------------------------------------ */
/* 命令实现                                                            */
/* ------------------------------------------------------------------ */

/// overrides 里的单元格 → 求值用的值（公式延后递归求值）
fn override_of(edit: &CellEdit) -> Override {
    match edit.kind.as_str() {
        "formula" => {
            let text = edit.value.trim().trim_start_matches('=').trim().to_string();
            if text.is_empty() {
                Override::Value(Value::error("公式为空"))
            } else {
                Override::Formula(text)
            }
        }
        "number" => Override::Value(match parse_number(&edit.value) {
            Some(n) => Value::Number(n),
            None => Value::error(ERR_VALUE),
        }),
        "bool" => Override::Value(match parse_bool(&edit.value) {
            Some(b) => Value::Bool(b),
            None => Value::error(ERR_VALUE),
        }),
        "empty" => Override::Value(Value::Empty),
        // text / date：按文本参与（日期不做序列号换算，见模块注释）
        _ => Override::Value(Value::Text(edit.value.clone())),
    }
}

/// 求值实现（同步，供命令在线程池里跑，也方便单元测试直接调用）
pub(crate) fn eval_impl(
    path: &str,
    sheet: &str,
    requests: &[EvalRequest],
    overrides: &[CellEdit],
) -> Result<Vec<EvalResult>, String> {
    let p = Path::new(path);
    let meta = std::fs::metadata(p).map_err(|e| format!("读取文件信息失败: {e}"))?;
    if !meta.is_file() {
        return Err("目标不是有效的文件".to_string());
    }
    // 影子优先 + 企业加密内存解密，全在 office.rs 这一条路径上
    let (bytes, _encrypted) = office::read_workbook_bytes(p)?;
    let workbook = office::parse_workbook(&bytes).map_err(|e| office::describe_parse_failure(&bytes, &e))?;
    let names = workbook.sheet_names();
    if !names.iter().any(|n| n.eq_ignore_ascii_case(sheet)) {
        return Err(format!("工作簿里没有工作表「{sheet}」"));
    }

    let mut map = HashMap::new();
    for edit in overrides {
        map.insert(
            (edit.sheet.to_lowercase(), edit.row, edit.col),
            override_of(edit),
        );
    }

    let mut book = Book {
        sheets: workbook,
        values: HashMap::new(),
        formulas: HashMap::new(),
        overrides: map,
        depth: 0,
        visiting: Vec::new(),
    };

    let mut out = Vec::with_capacity(requests.len());
    for req in requests {
        book.visiting.clear();
        book.depth = 0;
        // 把「本次要放进这个单元格的公式」临时登记成它自己的内容：这样公式引用自己所在的
        // 单元格（`=A1+1` 写在 A1）会被判成循环引用，而不是读到文件里的旧值。
        // 前端正常流程里这条编辑也在 overrides 里（那时不覆盖它，避免改变语义）。
        let key = (sheet.to_lowercase(), req.row, req.col);
        let installed = if book.overrides.contains_key(&key) {
            false
        } else {
            let text = req
                .formula
                .trim()
                .trim_start_matches('=')
                .trim()
                .to_string();
            if text.is_empty() {
                false
            } else {
                book.overrides.insert(key.clone(), Override::Formula(text));
                true
            }
        };
        let value = book
            .eval_at(sheet, req.row, req.col, &req.formula)
            // 顶层公式自身的错误也带上本格坐标（格式与区域传播一致，用户不用猜是哪一格）。
            // `locate` 是「已有位置优先」，所以区域里传上来的错误仍指向根因单元格。
            .locate(a1(req.row, req.col));
        if installed {
            book.overrides.remove(&key);
        }
        out.push(match value {
            Value::Error { .. } => EvalResult {
                row: req.row,
                col: req.col,
                value: None,
                error: Some(value.error_text()),
            },
            other => EvalResult {
                row: req.row,
                col: req.col,
                value: Some(other.display()),
                error: None,
            },
        });
    }
    Ok(out)
}

/// 求值一批公式（0.20.0「公式计算」）。
///
/// - `overrides` 是前端尚未落盘的单元格编辑（`CellEdit` 形状），求值时**优先**使用它们，
///   因此刚输入、还没保存的公式也能参与计算；其中若是公式则递归求值（深度上限 10）；
/// - 读数据走**影子优先**（结构改动后的坐标是对的），企业加密文件同样内存解密；
/// - 不支持的函数返回「暂不支持函数 XXX」，除零等返回 Excel 风格错误文本，
///   全部放在 `error` 字段里，`value` 与 `error` 互斥；
/// - 单条请求失败不影响其它请求（每条独立返回）。
#[tauri::command]
pub async fn spreadsheet_eval(
    path: String,
    sheet: String,
    requests: Vec<EvalRequest>,
    overrides: Vec<CellEdit>,
) -> Result<Vec<EvalResult>, String> {
    tauri::async_runtime::spawn_blocking(move || eval_impl(&path, &sheet, &requests, &overrides))
        .await
        .map_err(|e| format!("计算公式失败: {e}"))?
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::commands::esafenet;
    use rust_xlsxwriter::Workbook as XlsxWriter;
    use std::path::{Path as StdPath, PathBuf};

    fn test_dir(name: &str) -> PathBuf {
        let dir = std::env::temp_dir()
            .join("masteredit-office-formula")
            .join(name);
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).expect("创建测试目录");
        dir
    }

    /// 样本（0 起坐标）：
    /// 数据：A1=3 B1=4 C1="文本" E1=1 / A2=10 B2=20 E2=2 / A3="5"(数字文本) C3=TRUE / A4="abc"
    /// 第二表：A1=7，A2 是文件里的公式 =A1*3（缓存值为 0，用于验证「按公式重算」）
    fn make_sample(path: &StdPath) {
        let mut wb = XlsxWriter::new();
        {
            let s = wb.add_worksheet();
            s.set_name("数据").unwrap();
            s.write_number(0, 0, 3.0).unwrap(); // A1
            s.write_number(0, 1, 4.0).unwrap(); // B1
            s.write_string(0, 2, "文本").unwrap(); // C1
            s.write_number(0, 4, 1.0).unwrap(); // E1
            s.write_number(1, 0, 10.0).unwrap(); // A2
            s.write_number(1, 1, 20.0).unwrap(); // B2
            s.write_number(1, 4, 2.0).unwrap(); // E2
            s.write_string(2, 0, "5").unwrap(); // A3：数字文本
            s.write_boolean(2, 2, true).unwrap(); // C3：布尔
            s.write_string(3, 0, "abc").unwrap(); // A4
        }
        {
            let s = wb.add_worksheet();
            s.set_name("第二表").unwrap();
            s.write_number(0, 0, 7.0).unwrap();
            s.write_formula(1, 0, "=A1*3").unwrap(); // 文件里的公式
        }
        wb.save(path).expect("生成样本");
    }

    fn edit(sheet: &str, row: u32, col: u32, kind: &str, value: &str) -> CellEdit {
        CellEdit {
            sheet: sheet.to_string(),
            row,
            col,
            kind: kind.to_string(),
            value: value.to_string(),
        }
    }

    /// 造一个只有「数据」表的样本，单元格内容由调用方给定（0 起坐标）
    enum Spec<'a> {
        Num(f64),
        Text(&'a str),
        Formula(&'a str),
    }

    fn make_cells(path: &StdPath, cells: &[(u32, u16, Spec)]) {
        let mut wb = XlsxWriter::new();
        let s = wb.add_worksheet();
        s.set_name("数据").unwrap();
        for (row, col, spec) in cells {
            match spec {
                Spec::Num(n) => s.write_number(*row, *col, *n).unwrap(),
                Spec::Text(t) => s.write_string(*row, *col, *t).unwrap(),
                Spec::Formula(f) => s.write_formula(*row, *col, *f).unwrap(),
            };
        }
        wb.save(path).expect("生成样本");
    }

    fn request(row: u32, col: u32, formula: &str) -> EvalRequest {
        EvalRequest {
            row,
            col,
            formula: formula.to_string(),
        }
    }

    /// 求一条公式（可带 overrides），返回 (row, col, value, error)
    fn eval_at(
        path: &StdPath,
        row: u32,
        col: u32,
        formula: &str,
        overrides: &[CellEdit],
    ) -> (Option<String>, Option<String>) {
        let out = eval_impl(
            path.to_str().unwrap(),
            "数据",
            &[request(row, col, formula)],
            overrides,
        )
        .expect("求值应成功");
        assert_eq!(out.len(), 1);
        (out[0].value.clone(), out[0].error.clone())
    }

    /// 求值并断言成功，返回显示文本
    fn value(path: &StdPath, formula: &str) -> String {
        let (v, e) = eval_at(path, 20, 0, formula, &[]);
        assert!(e.is_none(), "「{formula}」不应报错：{e:?}");
        v.unwrap()
    }

    /// 求值并断言失败，返回错误文本
    fn error(path: &StdPath, formula: &str) -> String {
        let (v, e) = eval_at(path, 20, 0, formula, &[]);
        assert!(v.is_none(), "「{formula}」不应算出值：{v:?}");
        e.unwrap()
    }

    /* ------------------------------ 基础 ------------------------------ */

    #[test]
    fn sums_ranges_whole_column_and_whole_row() {
        let dir = test_dir("sum");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.as_path();

        assert_eq!(value(p, "=SUM(E1:E2)"), "3");
        assert_eq!(value(p, "=sum(e1:e2)"), "3", "函数名与引用不区分大小写");
        assert_eq!(value(p, "=SUM(E:E)"), "3", "整列");
        assert_eq!(value(p, "=SUM(2:2)"), "32", "整行：A2(10)+B2(20)+E2(2)");
        // A 列：3 + 10 + "5"(数字文本) + "abc"(忽略) = 18
        assert_eq!(value(p, "=SUM(A:A)"), "18", "数字文本参与、普通文本忽略");
        assert_eq!(value(p, "=SUM(A1,B1)"), "7", "多个参数");
        assert_eq!(value(p, "=SUM(A1:B2,5)"), "42", "区域 + 标量");
    }

    #[test]
    fn arithmetic_comparison_and_percent() {
        let dir = test_dir("math");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.as_path();

        assert_eq!(value(p, "=A1+B1*2"), "11");
        assert_eq!(value(p, "=(A1+B1)/2"), "3.5");
        assert_eq!(value(p, "=A1^2"), "9");
        assert_eq!(value(p, "=50%"), "0.5");
        assert_eq!(value(p, "=-A1"), "-3");
        assert_eq!(value(p, "=A1>B1"), "FALSE");
        assert_eq!(value(p, "=A1<B1"), "TRUE");
        assert_eq!(value(p, "=A1<>B1"), "TRUE");
        assert_eq!(value(p, "=A1>=3"), "TRUE");
        assert_eq!(value(p, "=\"abc\"=\"ABC\""), "TRUE", "文本比较不区分大小写");
        assert_eq!(value(p, "=B1&\"-\"&A1"), "4-3", "& 连接");
    }

    #[test]
    fn aggregates_on_mixed_text_and_empty_range() {
        let dir = test_dir("aggregate");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.as_path();

        // A1:C3 = 3,4,文本 / 10,20,空 / "5",空,TRUE
        assert_eq!(value(p, "=SUM(A1:C3)"), "42", "数字文本参与、文本/布尔/空忽略");
        assert_eq!(value(p, "=COUNT(A1:C3)"), "5", "COUNT 只数数字");
        assert_eq!(
            value(p, "=COUNTA(A1:C3)"),
            "7",
            "COUNTA 数非空（3,4,文本,10,20,5,TRUE）"
        );
        assert_eq!(value(p, "=MIN(A1:C3)"), "3");
        assert_eq!(value(p, "=MAX(A1:C3)"), "20");
        assert_eq!(value(p, "=AVERAGE(A1:C3)"), "8.4");
        // 全文本区域：SUM 0、AVERAGE 报 #DIV/0!
        assert_eq!(value(p, "=SUM(A4)"), "0");
        assert_eq!(
            error(p, "=AVERAGE(A4)"),
            "#DIV/0!（A21 出错）",
            "全文本区域求平均：错误出在公式自身（请求格 A21）"
        );
        assert_eq!(value(p, "=MIN(A4)"), "0", "没有数字时 MIN 给 0（与 Excel 一致）");
    }

    #[test]
    fn if_and_iferror() {
        let dir = test_dir("if");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.as_path();

        assert_eq!(value(p, "=IF(A1>1,\"大\",\"小\")"), "大");
        assert_eq!(value(p, "=IF(A1<1,\"大\",\"小\")"), "小");
        assert_eq!(value(p, "=IF(A1<1,\"大\")"), "FALSE", "省略 else 时返回 FALSE");
        assert_eq!(value(p, "=IFERROR(1/0,\"err\")"), "err");
        assert_eq!(value(p, "=IFERROR(A1,\"err\")"), "3");
        // IF 只算被选中的分支：没选中的分支报错也不影响
        assert_eq!(value(p, "=IF(TRUE,1,1/0)"), "1");
        assert_eq!(value(p, "=AND(A1>0,B1>0)"), "TRUE");
        assert_eq!(value(p, "=OR(A1>100,B1>0)"), "TRUE");
        assert_eq!(value(p, "=NOT(A1>0)"), "FALSE");
    }

    #[test]
    fn text_and_math_functions() {
        let dir = test_dir("functions");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.as_path();

        assert_eq!(value(p, "=ROUND(1.234,2)"), "1.23");
        assert_eq!(value(p, "=ROUND(2.5,0)"), "3", "Excel 的 ROUND 是四舍五入（远离 0）");
        assert_eq!(value(p, "=ROUND(1.5)"), "2", "省略位数按 0 位");
        assert_eq!(value(p, "=ABS(-4)"), "4");
        assert_eq!(value(p, "=INT(2.9)"), "2");
        assert_eq!(value(p, "=LEN(\"abc\")"), "3");
        assert_eq!(value(p, "=LEN(\"中文\")"), "2", "按字符数");
        assert_eq!(value(p, "=LEFT(\"abcd\",2)"), "ab");
        assert_eq!(value(p, "=RIGHT(\"abcd\",2)"), "cd");
        assert_eq!(value(p, "=MID(\"abcdef\",2,3)"), "bcd");
        assert_eq!(value(p, "=TRIM(\"  a   b  \")"), "a b");
        assert_eq!(value(p, "=UPPER(\"ab\")"), "AB");
        assert_eq!(value(p, "=LOWER(\"AB\")"), "ab");
        assert_eq!(value(p, "=CONCAT(A1,\"-\",B1)"), "3-4");
        assert_eq!(value(p, "=CONCATENATE(\"a\",\"b\")"), "ab");
        assert_eq!(value(p, "=LEN(A1)"), "1", "数字按显示文本参与");
    }

    #[test]
    fn unsupported_function_reports_clearly() {
        let dir = test_dir("unsupported");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.as_path();

        let err = error(p, "=VLOOKUP(A1,A1:B3,2,FALSE)");
        assert!(err.contains("暂不支持函数"), "实际提示：{err}");
        assert!(err.contains("VLOOKUP"), "应带上函数名：{err}");
        assert_eq!(err, "暂不支持函数 VLOOKUP（A21）", "顶层错误也带本格坐标");
        assert!(error(p, "=SUMIF(A1:A3,\">1\")").contains("暂不支持函数 SUMIF"));
        // 语法错误也不能 panic
        assert!(error(p, "=SUM(").contains("语法错误"));
        assert!(error(p, "=1+").contains("语法错误"));
        assert!(error(p, "").contains("公式为空"));
        // Excel 风格错误（位置就是写出这条公式的单元格）
        assert_eq!(error(p, "=1/0"), "#DIV/0!（A21 出错）");
        assert_eq!(error(p, "=A4+1"), "#VALUE!（A21 出错）");
        assert_eq!(error(p, "=没有这个表!A1"), "#REF!（A21 出错）");
    }

    /// 顶层公式自身的错误带本格坐标；区域里传上来的错误保留根因坐标（不被覆盖）
    #[test]
    fn top_level_errors_carry_their_own_position() {
        let dir = test_dir("err-top-level");
        let path = dir.join("a.xlsx");
        make_cells(
            &path,
            &[
                (4, 2, Spec::Formula("=VLOOKUP(1,2,3)")), // C5：用不支持函数的公式
                (4, 3, Spec::Text("abc")),                 // D5：文本（供 #VALUE! 用）
            ],
        );
        let p = path.as_path();

        // B5 自己写错 → 位置就是 B5
        let (v, e) = eval_at(p, 4, 1, "=VLOOKUP(1,2,3)", &[]);
        assert!(v.is_none());
        assert_eq!(e.as_deref(), Some("暂不支持函数 VLOOKUP（B5）"));
        let (v, e) = eval_at(p, 4, 1, "=1/0", &[]);
        assert!(v.is_none());
        assert_eq!(e.as_deref(), Some("#DIV/0!（B5 出错）"));
        let (v, e) = eval_at(p, 4, 1, "=\"abc\"+1", &[]);
        assert!(v.is_none());
        assert_eq!(e.as_deref(), Some("#VALUE!（B5 出错）"));

        // 引用了文本格参与算术：错误发生在本格这条公式上（与「在 B5 写 =A1+1 且 A1 是文本」一致）
        let (v, e) = eval_at(p, 4, 1, "=D5+1", &[]);
        assert!(v.is_none());
        assert_eq!(e.as_deref(), Some("#VALUE!（B5 出错）"));
        // 区域里传上来的错误同样保留根因（C5），不会被请求格（B5）覆盖
        let (v, e) = eval_at(p, 4, 1, "=SUM(C5:C5)", &[]);
        assert!(v.is_none());
        assert_eq!(e.as_deref(), Some("暂不支持函数 VLOOKUP（C5）"));
        let (v, e) = eval_at(p, 4, 1, "=C5+1", &[]);
        assert!(v.is_none());
        assert_eq!(e.as_deref(), Some("暂不支持函数 VLOOKUP（C5）"));
    }

    #[test]
    fn today_and_now_are_date_text() {
        let dir = test_dir("today");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.as_path();

        let today = value(p, "=TODAY()");
        assert_eq!(today.len(), 10, "YYYY-MM-DD：{today}");
        assert_eq!(&today[4..5], "-");
        assert_eq!(&today[7..8], "-");
        assert!(today.chars().filter(|c| c.is_ascii_digit()).count() == 8);

        let now = value(p, "=NOW()");
        assert_eq!(now.len(), 19, "YYYY-MM-DD HH:MM:SS：{now}");
        assert!(now.contains(' ') && now.matches(':').count() == 2);
    }

    /* --------------------------- overrides --------------------------- */

    #[test]
    fn overrides_win_over_file_and_recurse() {
        let dir = test_dir("overrides");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.as_path();

        // 磁盘上 A1=3；override 改成 5 → SUM(A1:A2) = 5 + 10
        let ov = [edit("数据", 0, 0, "number", "5")];
        let (v, e) = eval_at(p, 20, 0, "=SUM(A1:A2)", &ov);
        assert_eq!(e, None);
        assert_eq!(v.as_deref(), Some("15"));

        // override 本身是公式 → 递归求值
        let ov = [edit("数据", 0, 0, "formula", "=10+5")];
        assert_eq!(eval_at(p, 20, 0, "=SUM(A1:A2)", &ov).0.as_deref(), Some("25"));

        // 链式：A1 = B1*2、B1 = 3+1 → 8
        let ov = [
            edit("数据", 0, 0, "formula", "=B1*2"),
            edit("数据", 0, 1, "formula", "=3+1"),
        ];
        assert_eq!(eval_at(p, 20, 0, "=A1", &ov).0.as_deref(), Some("8"));

        // 文本 / 布尔 / 清空 编辑
        let ov = [edit("数据", 0, 0, "text", "abc")];
        assert_eq!(eval_at(p, 20, 0, "=A1", &ov).0.as_deref(), Some("abc"));
        assert_eq!(
            eval_at(p, 20, 0, "=SUM(A1:A2)", &ov).0.as_deref(),
            Some("10"),
            "文本不参与 SUM"
        );
        let ov = [edit("数据", 0, 0, "bool", "true")];
        assert_eq!(eval_at(p, 20, 0, "=A1", &ov).0.as_deref(), Some("TRUE"));
        let ov = [edit("数据", 0, 0, "empty", "")];
        assert_eq!(eval_at(p, 20, 0, "=SUM(A1:A2)", &ov).0.as_deref(), Some("10"));
    }

    #[test]
    fn file_formulas_are_recomputed_and_cross_sheet_works() {
        let dir = test_dir("cross-sheet");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.as_path();

        assert_eq!(value(p, "=第二表!A1"), "7", "跨表引用");
        // 第二表!A2 是文件里的公式 =A1*3，缓存值是 0：必须重算成 21
        assert_eq!(value(p, "=第二表!A2"), "21");
        assert_eq!(value(p, "=第二表!A1+第二表!A2"), "28");
        assert_eq!(value(p, "=SUM(第二表!A1:A2)"), "28");
    }

    /* ------------------------ 循环引用与深度 ------------------------ */

    #[test]
    fn circular_reference_is_reported_without_hanging() {
        let dir = test_dir("circular");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.as_path();

        // 请求自己引用自己（请求所在单元格是 A21，所以这里直接用 A1 引用 A1 的覆盖）
        let ov = [edit("数据", 0, 0, "formula", "=A1+1")];
        assert_eq!(eval_at(p, 20, 0, "=A1", &ov).1.as_deref(), Some("循环引用（A1）"));

        // 互相引用
        let ov = [
            edit("数据", 0, 0, "formula", "=B1"),
            edit("数据", 0, 1, "formula", "=A1"),
        ];
        assert_eq!(eval_at(p, 20, 0, "=A1", &ov).1.as_deref(), Some("循环引用（A1）"));

        // 请求单元格自身被引用（=A21 写在 A21）
        let (v, e) = eval_at(p, 20, 0, "=A21+1", &[]);
        assert!(v.is_none());
        assert_eq!(e.as_deref(), Some("循环引用（A21）"));
    }

    #[test]
    fn depth_limit_is_reported() {
        let dir = test_dir("depth");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.as_path();

        // A1 = A2, A2 = A3, ... A11 = 1 → 链太长
        let mut ov = Vec::new();
        for i in 0..10u32 {
            ov.push(edit(
                "数据",
                i,
                3,
                "formula",
                &format!("=D{}", i + 2),
            ));
        }
        ov.push(edit("数据", 10, 3, "number", "1"));
        let (v, e) = eval_at(p, 20, 0, "=D1", &ov);
        assert!(v.is_none(), "不该算出值：{v:?}");
        assert!(
            e.as_deref().unwrap_or_default().starts_with("公式引用层级过深"),
            "实际提示：{e:?}"
        );
    }

    /* --------------------------- 集成路径 --------------------------- */

    #[test]
    fn encrypted_file_is_decrypted_in_memory() {
        let dir = test_dir("encrypted");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let plain = std::fs::read(&path).unwrap();
        let mut header = vec![0u8; 4096];
        header[0..4].copy_from_slice(&[0xE0, 0xA8, 0x91, 0xE7]);
        header[8..12].copy_from_slice(&512u32.to_le_bytes());
        header[12..16].copy_from_slice(&4096u32.to_le_bytes());
        std::fs::write(&path, esafenet::encrypt_esafenet(&header, &plain)).unwrap();

        assert_eq!(value(path.as_path(), "=SUM(E1:E2)"), "3");
        assert_eq!(value(path.as_path(), "=第二表!A2"), "21");
    }

    /// 结构操作改了影子之后，求值用的是**影子内容**（引用跟着影子走）
    #[test]
    fn eval_follows_shadow_after_structure_change() {
        use crate::commands::office_ops::StructureOp;
        use crate::commands::office_shadow::{self, SHADOWS};

        let dir = test_dir("shadow");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let p = path.to_str().unwrap().to_string();

        assert_eq!(value(path.as_path(), "=SUM(E1:E2)"), "3");

        // 在第 1 行之前插入一行：E1 变空、E2 变成原来的 E1（1）
        office_shadow::apply_op(
            &SHADOWS,
            &p,
            &StructureOp::InsertRows {
                sheet: "数据".to_string(),
                at: 0,
                count: 1,
            },
            &[],
        )
        .expect("插入行应成功");
        assert_eq!(
            value(path.as_path(), "=SUM(E1:E2)"),
            "1",
            "求值必须用影子内容（磁盘上还是 3）"
        );
        // 磁盘内容没变：读磁盘仍是 3
        let disk_only = {
            let bytes = std::fs::read(&path).unwrap();
            let mut wb = calamine::open_workbook_auto_from_rs(Cursor::new(bytes)).unwrap();
            wb.worksheet_range("数据")
                .unwrap()
                .get_value((1, 4))
                .cloned()
        };
        assert_eq!(disk_only, Some(Data::Float(2.0)), "磁盘上 E2 仍是 2");

        // 丢弃影子后又回到磁盘内容
        assert!(office_shadow::discard(&p));
        assert_eq!(value(path.as_path(), "=SUM(E1:E2)"), "3");
    }

    #[test]
    fn missing_sheet_is_reported() {
        let dir = test_dir("missing");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let err = eval_impl(
            path.to_str().unwrap(),
            "不存在的表",
            &[request(0, 0, "=1+1")],
            &[],
        )
        .expect_err("表不存在应报错");
        assert!(err.contains("不存在的表"), "实际提示：{err}");
    }

    /// 一批请求里单条失败不影响其它条
    #[test]
    fn batch_results_are_independent() {
        let dir = test_dir("batch");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let out = eval_impl(
            path.to_str().unwrap(),
            "数据",
            &[
                request(20, 0, "=1+1"),
                request(20, 1, "=VLOOKUP(1,2,3)"),
                request(20, 2, "=A1"),
            ],
            &[],
        )
        .unwrap();
        assert_eq!(out.len(), 3);
        assert_eq!(out[0].value.as_deref(), Some("2"));
        assert!(out[1].error.as_deref().unwrap().contains("暂不支持"));
        assert_eq!(out[2].value.as_deref(), Some("3"));
        assert_eq!((out[0].row, out[0].col), (20, 0));
        // value / error 互斥
        for item in &out {
            assert!(item.value.is_some() != item.error.is_some(), "{item:?}");
        }
    }

    /* --------------------------- 错误传播 --------------------------- */

    /// 要求 1/3：区域里的**普通文本仍然忽略**（SUM 跳过、COUNT 只数数字、COUNTA 数非空）
    #[test]
    fn aggregates_still_ignore_plain_text() {
        let dir = test_dir("err-text");
        let path = dir.join("a.xlsx");
        make_cells(
            &path,
            &[
                (0, 0, Spec::Num(1.0)),
                (1, 0, Spec::Text("abc")),
                (2, 0, Spec::Num(3.0)),
            ],
        );
        let p = path.as_path();
        assert_eq!(value(p, "=SUM(A1:A3)"), "4");
        assert_eq!(value(p, "=AVERAGE(A1:A3)"), "2");
        assert_eq!(value(p, "=COUNT(A1:A3)"), "2", "只数数字");
        assert_eq!(value(p, "=COUNTA(A1:A3)"), "3", "非空都算");
        assert_eq!(value(p, "=MIN(A1:A3)"), "1");
        assert_eq!(value(p, "=MAX(A1:A3)"), "3");
    }

    /// 要求 1/5：区域里某格是**求值后出错**的公式 → 聚合结果就是那个错误（带位置）
    #[test]
    fn errors_in_range_propagate_with_position() {
        let dir = test_dir("err-propagate");
        let path = dir.join("a.xlsx");
        make_cells(
            &path,
            &[
                (0, 0, Spec::Num(1.0)),
                (1, 0, Spec::Formula("=1/0")), // A2 求值得到 #DIV/0!
                (2, 0, Spec::Num(3.0)),
            ],
        );
        let p = path.as_path();
        assert_eq!(error(p, "=SUM(A1:A3)"), "#DIV/0!（A2 出错）");
        assert_eq!(
            error(p, "=SUM(1,2/0)"),
            "#DIV/0!（A21 出错）",
            "错误在公式自身（不是引用来的）→ 位置是本格"
        );
    }

    /// 要求 3（测试清单第 3 条）：所有聚合都要传播，而不是只传播 SUM
    #[test]
    fn all_aggregates_propagate_errors() {
        let dir = test_dir("err-all");
        let path = dir.join("a.xlsx");
        make_cells(
            &path,
            &[
                (0, 0, Spec::Num(1.0)),
                (1, 0, Spec::Formula("=1/0")),
                (2, 0, Spec::Num(3.0)),
            ],
        );
        let p = path.as_path();
        for f in [
            "=SUM(A1:A3)",
            "=AVERAGE(A1:A3)",
            "=COUNT(A1:A3)",
            "=COUNTA(A1:A3)",
            "=MIN(A1:A3)",
            "=MAX(A1:A3)",
        ] {
            let (v, e) = eval_at(p, 20, 0, f, &[]);
            assert!(v.is_none(), "「{f}」不该算出数字：{v:?}");
            assert_eq!(e.as_deref(), Some("#DIV/0!（A2 出错）"), "「{f}」应传播错误");
        }
    }

    /// 要求 2/5：区域里某格的公式用了不支持的函数 → 整条链都传播，且带位置
    #[test]
    fn unsupported_function_in_range_propagates() {
        let dir = test_dir("err-unsupported");
        let path = dir.join("a.xlsx");
        make_cells(
            &path,
            &[
                (0, 0, Spec::Num(1.0)),
                (1, 0, Spec::Formula("=VLOOKUP(A1,A1:B3,2,FALSE)")),
                (2, 0, Spec::Num(3.0)),
            ],
        );
        let p = path.as_path();
        let err = error(p, "=SUM(A1:A3)");
        assert!(err.contains("暂不支持函数 VLOOKUP"), "实际提示：{err}");
        assert!(err.ends_with("（A2）"), "中文提示的位置格式：{err}");
        assert_eq!(err, "暂不支持函数 VLOOKUP（A2）");
        // 间接引用也一样：B1 = SUM(A1:A3) → 根因仍是 A2
        let nested = dir.join("b.xlsx");
        make_cells(
            &nested,
            &[
                (0, 0, Spec::Num(1.0)),
                (1, 0, Spec::Formula("=VLOOKUP(A1,A1:B3,2,FALSE)")),
                (2, 0, Spec::Num(3.0)),
                (0, 1, Spec::Formula("=SUM(A1:A3)")), // B1
            ],
        );
        assert_eq!(
            eval_at(nested.as_path(), 20, 0, "=B1", &[]).1.as_deref(),
            Some("暂不支持函数 VLOOKUP（A2）"),
            "嵌套引用要保留根因位置"
        );
    }

    /// 要求（测试清单第 5 条）：区域外的错误不影响结果
    #[test]
    fn errors_outside_range_do_not_affect() {
        let dir = test_dir("err-outside");
        let path = dir.join("a.xlsx");
        make_cells(
            &path,
            &[
                (0, 0, Spec::Num(1.0)),
                (1, 0, Spec::Num(2.0)),
                (2, 0, Spec::Num(3.0)),
                (4, 0, Spec::Formula("=1/0")), // A5 出错，但在 A1:A3 之外
            ],
        );
        let p = path.as_path();
        assert_eq!(value(p, "=SUM(A1:A3)"), "6");
        assert_eq!(value(p, "=MAX(A1:A3)"), "3");
        assert_eq!(error(p, "=SUM(A1:A5)"), "#DIV/0!（A5 出错）", "扩到 A5 就会传播");
    }

    /// 要求 6（测试清单第 6 条）：多处错误取先行后列的第一个
    #[test]
    fn first_error_in_row_major_order_wins() {
        let dir = test_dir("err-first");
        let path = dir.join("a.xlsx");
        make_cells(
            &path,
            &[
                (0, 0, Spec::Num(1.0)),
                (0, 1, Spec::Formula("=1/0")), // B1：行序更靠前
                (1, 0, Spec::Formula("=VLOOKUP(1,2,3)")), // A2
                (1, 1, Spec::Num(2.0)),
            ],
        );
        let p = path.as_path();
        assert_eq!(error(p, "=SUM(A1:B2)"), "#DIV/0!（B1 出错）");
        // 反过来：把错误放到 A2、B1 换成正常值，则应指向 A2
        let other = dir.join("b.xlsx");
        make_cells(
            &other,
            &[
                (0, 0, Spec::Num(1.0)),
                (0, 1, Spec::Num(2.0)),
                (1, 0, Spec::Formula("=1/0")),
                (1, 1, Spec::Num(3.0)),
            ],
        );
        assert_eq!(
            eval_at(other.as_path(), 20, 0, "=SUM(A1:B2)", &[]).1.as_deref(),
            Some("#DIV/0!（A2 出错）")
        );
    }

    /// 要求（测试清单第 7 条）：整列引用里的错误同样传播
    #[test]
    fn whole_column_reference_propagates_errors() {
        let dir = test_dir("err-column");
        let path = dir.join("a.xlsx");
        make_cells(
            &path,
            &[
                (0, 1, Spec::Num(1.0)),          // B1
                (1, 1, Spec::Formula("=1/0")),   // B2
                (2, 1, Spec::Num(3.0)),          // B3
            ],
        );
        let p = path.as_path();
        assert_eq!(error(p, "=SUM(B:B)"), "#DIV/0!（B2 出错）");
        // 整行同理
        let row_err = dir.join("b.xlsx");
        make_cells(
            &row_err,
            &[
                (1, 0, Spec::Num(1.0)),        // A2
                (1, 1, Spec::Formula("=1/0")), // B2
            ],
        );
        assert_eq!(
            eval_at(row_err.as_path(), 20, 0, "=SUM(2:2)", &[]).1.as_deref(),
            Some("#DIV/0!（B2 出错）")
        );
    }

    /// 要求（测试清单第 8 条）：IFERROR 能捕获被传播的错误（错误走普通错误通道）
    #[test]
    fn iferror_catches_propagated_error() {
        let dir = test_dir("err-iferror");
        let path = dir.join("a.xlsx");
        make_cells(
            &path,
            &[
                (0, 0, Spec::Num(1.0)),
                (1, 0, Spec::Formula("=1/0")),
                (2, 0, Spec::Num(3.0)),
            ],
        );
        let p = path.as_path();
        assert_eq!(value(p, "=IFERROR(SUM(A1:A3),\"x\")"), "x");
        assert_eq!(value(p, "=IFERROR(AVERAGE(A1:A3),\"没算出来\")"), "没算出来");
        assert_eq!(value(p, "=IFERROR(SUM(A1:A1),\"x\")"), "1", "没出错时返回原值");
        assert_eq!(
            error(p, "=SUM(A1:A3)+1"),
            "#DIV/0!（A2 出错）",
            "错误参与算术运算时照样传播（带位置）"
        );
    }

    /// 文件里**缓存**为错误值的单元格（不是公式，直接是 &lt;v&gt;#DIV/0!&lt;/v&gt;）同样传播
    #[test]
    fn cached_error_value_propagates() {
        let dir = test_dir("err-cached");
        let path = dir.join("a.xlsx");
        {
            let mut wb = umya_spreadsheet::new_file();
            let s = wb.sheet_by_name_mut("Sheet1").unwrap();
            s.set_name("数据");
            s.cell_mut((1, 1)).set_value_number(1.0); // A1
            s.cell_mut((1, 2)).set_error("#DIV/0!"); // A2：缓存错误值
            s.cell_mut((1, 3)).set_value_number(3.0); // A3
            umya_spreadsheet::writer::xlsx::write_writer(
                &wb,
                &mut std::fs::File::create(&path).unwrap(),
            )
            .unwrap();
        }
        let p = path.as_path();
        assert_eq!(error(p, "=SUM(A1:A3)"), "#DIV/0!（A2 出错）");
        assert_eq!(error(p, "=A2+1"), "#DIV/0!（A2 出错）", "单格引用也带位置");
    }

    /* --------------------------- 前后端契约 --------------------------- */

    #[test]
    fn serde_wire_format_matches_frontend_contract() {
        let req: EvalRequest =
            serde_json::from_str(r#"{"row":1,"col":2,"formula":"=SUM(A1:A2)"}"#).unwrap();
        assert_eq!(req, request(1, 2, "=SUM(A1:A2)"));

        let ok = EvalResult {
            row: 1,
            col: 2,
            value: Some("3".to_string()),
            error: None,
        };
        assert_eq!(
            serde_json::to_string(&ok).unwrap(),
            r#"{"row":1,"col":2,"value":"3","error":null}"#
        );
        let bad = EvalResult {
            row: 0,
            col: 0,
            value: None,
            error: Some("暂不支持函数 VLOOKUP".to_string()),
        };
        assert_eq!(
            serde_json::to_string(&bad).unwrap(),
            r#"{"row":0,"col":0,"value":null,"error":"暂不支持函数 VLOOKUP"}"#
        );
    }
}
