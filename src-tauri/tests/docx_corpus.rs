//! DOCX 语料库完备性测试：拿**真实的、批量加密的** Word 文档跑一遍解析。
//!
//! 这是 `docs/plan-docx.md` 里的验收红线：
//! **每个文件要么解析成功，要么给出可操作的中文错误，绝不能崩、不能静默失败。**
//!
//! 语料库目录不存在时**自动跳过**（其它机器 / CI 上没有这个目录，不应阻塞）。
//! 用 `MASTEREDIT_DOCX_CORPUS` 指定别的目录。
//!
//! 另外补了两个**合成样本**（见文件末尾）：真实语料覆盖不到「只写字符单位缩进」与
//! 「自动段间距」这两条路径（扫过 25 个真实文档：Chars 只有 1 个且是"两个都写"的形态、
//! Autospacing 一个都没有），合成样本用来挡住这两条路径的回归。

use std::io::Write;
use std::path::{Path, PathBuf};

use masteredit_lib::commands::office_docx::{document_blocks, document_info, document_xml};

fn corpus_dir() -> PathBuf {
    match std::env::var("MASTEREDIT_DOCX_CORPUS") {
        Ok(value) => PathBuf::from(value),
        Err(_) => PathBuf::from(r"Z:\D\mywork\05_会议汇报\每周汇报\2026"),
    }
}

/// 递归收集目录下的 .docx（按路径排序，保证输出稳定）
///
/// 跳过 Word 的临时锁文件 `~$xxx.docx`：那是 Word 打开文档时生成的 owner 文件
/// （记录"谁打开了它"），不是文档本身、也不是 zip，扫进来只会制造假失败。
fn collect_docx(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_docx(&path, out);
        } else if path
            .extension()
            .is_some_and(|ext| ext.eq_ignore_ascii_case("docx"))
            && !path
                .file_name()
                .is_some_and(|name| name.to_string_lossy().starts_with("~$"))
        {
            out.push(path);
        }
    }
    out.sort();
}

#[test]
fn every_corpus_file_parses_or_fails_cleanly() {
    let dir = corpus_dir();
    if !dir.exists() {
        eprintln!("跳过：语料库目录不存在 {}", dir.display());
        return;
    }
    let mut files = Vec::new();
    collect_docx(&dir, &mut files);
    if files.is_empty() {
        eprintln!("跳过：目录里没有 .docx —— {}", dir.display());
        return;
    }

    let mut ok_count = 0usize;
    let mut encrypted_count = 0usize;
    let mut total_paragraphs = 0usize;
    let mut total_tables = 0usize;
    let mut failures: Vec<String> = Vec::new();

    for path in &files {
        let label = path.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default();
        match document_info(path.to_string_lossy().to_string()) {
            Ok(info) => {
                ok_count += 1;
                if info.encrypted {
                    encrypted_count += 1;
                }
                total_paragraphs += info.paragraphs;
                total_tables += info.tables;
                // 结构自检：必须是合法 docx（有 document.xml），且只读定位成立
                assert!(
                    info.parts.iter().any(|part| part.name == "word/document.xml"),
                    "{label}: 解析成功但包里没有 word/document.xml"
                );
                assert!(!info.editable, "{label}: DOCX 应当一律只读");
                // 正文能读出来（空文档除外）
                if info.paragraphs > 0 {
                    let xml = document_xml(path.to_string_lossy().to_string())
                        .unwrap_or_else(|error| panic!("{label}: 读 document.xml 失败：{error}"));
                    assert!(xml.starts_with("<?xml"), "{label}: document.xml 不是 XML 文本");
                    assert!(xml.contains("<w:body"), "{label}: document.xml 缺少文档主体");
                }
                eprintln!(
                    "  [OK]   {label:<28} 部件 {:>3} · 段落 {:>4} · 表格 {:>2} · 图片 {:>2}{}",
                    info.parts.len(),
                    info.paragraphs,
                    info.tables,
                    info.images,
                    if info.encrypted { " · 已解密" } else { "" }
                );
            }
            Err(message) => {
                // 失败也必须"可操作"：不能是底层 zip/panic 噪音，要能照着做
                let actionable = message.contains("不是有效的 .docx")
                    || message.contains(".doc")
                    || message.contains("密码")
                    || message.contains("文件不存在")
                    || message.contains("读取文件失败");
                assert!(
                    actionable,
                    "{label}: 错误信息不够可操作，需要给出下一步建议：{message}"
                );
                failures.push(format!("{label}: {message}"));
                eprintln!("  [FAIL] {label}: {message}");
            }
        }
    }

    eprintln!(
        "\n语料库统计：共 {} 个文件，成功 {} 个（其中加密 {} 个），失败 {} 个；累计段落 {}、表格 {}",
        files.len(),
        ok_count,
        encrypted_count,
        failures.len(),
        total_paragraphs,
        total_tables
    );

    assert!(
        failures.is_empty(),
        "有 {} 个文件解析失败：\n{}",
        failures.len(),
        failures.join("\n")
    );
    assert_eq!(ok_count, files.len(), "应当每个文件都能解析");
}

/// 语料库里每个文件的 `document.xml` 都应当是**结构完整**的 XML：
/// 首尾标签闭合、没有截断（加密解密链路一旦出错，这里最先暴露）
#[test]
fn corpus_document_xml_is_well_formed() {
    let dir = corpus_dir();
    if !dir.exists() {
        eprintln!("跳过：语料库目录不存在 {}", dir.display());
        return;
    }
    let mut files = Vec::new();
    collect_docx(&dir, &mut files);
    if files.is_empty() {
        return;
    }
    let mut checked = 0usize;
    for path in files.iter().take(30) {
        let Ok(xml) = document_xml(path.to_string_lossy().to_string()) else {
            continue;
        };
        let mut local = quick_xml::Reader::from_str(&xml);
        local.config_mut().trim_text(false);
        let mut depth = 0i32;
        let mut buffer = Vec::new();
        loop {
            match local.read_event_into(&mut buffer) {
                Ok(quick_xml::events::Event::Start(_)) => depth += 1,
                Ok(quick_xml::events::Event::End(_)) => depth -= 1,
                Ok(quick_xml::events::Event::Eof) => break,
                Ok(_) => {}
                Err(error) => panic!(
                    "{}: document.xml 在偏移 {} 处解析失败：{error}",
                    path.display(),
                    local.buffer_position()
                ),
            }
            buffer.clear();
        }
        assert_eq!(depth, 0, "{}: document.xml 标签未闭合（深度 {depth}）", path.display());
        checked += 1;
    }
    eprintln!("XML 结构校验：{checked} 个文件的 document.xml 全部良构");
}

/* ------------------------- 合成样本（补真实语料的盲区） ------------------------- */

/// 造一个最小 docx：只含 `word/document.xml`（解析器不需要其它部件也能工作）
fn write_minimal_docx(path: &Path, document_xml: &str) {
    let file = std::fs::File::create(path).expect("创建样本失败");
    let mut writer = zip::ZipWriter::new(file);
    let options: zip::write::FileOptions<'_, ()> =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Stored);
    writer.start_file("word/document.xml", options).expect("写入部件失败");
    writer.write_all(document_xml.as_bytes()).expect("写入失败");
    writer.finish().expect("收尾失败");
}

/// 取第一个段落块（合成样本只有一个段落）
fn first_paragraph(path: &Path) -> serde_json::Value {
    let page = document_blocks(path.to_string_lossy().to_string(), 0, 10).expect("取块失败");
    let blocks = serde_json::to_value(&page.blocks).expect("序列化失败");
    blocks
        .as_array()
        .and_then(|list| list.iter().find(|block| block["kind"] == "paragraph").cloned())
        .expect("样本里应有段落块")
}

/// 真实语料盖不到的两条路径：**只写字符单位缩进**、**自动段间距**。
/// 扫过 25 个真实文档：Chars 只有 1 个且是"两个都写"的形态、Autospacing 一个都没有 ——
/// 这两条只靠合成样本挡回归。
#[test]
fn synthetic_edge_cases_resolve_like_word() {
    let dir = std::env::temp_dir().join("masteredit-docx-synthetic");
    std::fs::create_dir_all(&dir).expect("创建临时目录失败");

    // 1) 只写 firstLineChars="200"（首行缩进 2 字符）不给 twips：按最终字号折算
    let chars_sample = dir.join("chars-indent.docx");
    write_minimal_docx(
        &chars_sample,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:pPr><w:ind w:firstLineChars="200"/></w:pPr><w:r><w:t>字符单位缩进</w:t></w:r></w:p>
</w:body></w:document>"#,
    );
    let paragraph = first_paragraph(&chars_sample);
    let indent = paragraph["indentFirstLinePt"].as_f64().unwrap_or_default();
    assert!(
        (indent - 21.0).abs() < 0.51,
        "只写 firstLineChars=200 时应按最终字号折算（默认 10.5pt -> 21pt），实际 {indent}"
    );
    assert_eq!(
        paragraph["text"].as_str().unwrap_or_default(),
        "字符单位缩进",
        "正文应能读出来"
    );

    // 2) 自动段间距：docDefaults 写死 before=0 after=0，开关打开后应取 14pt（不被 0 吃掉）
    let auto_sample = dir.join("auto-spacing.docx");
    write_minimal_docx(
        &auto_sample,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:pPr><w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr><w:r><w:t>自动段间距</w:t></w:r></w:p>
</w:body></w:document>"#,
    );
    let paragraph = first_paragraph(&auto_sample);
    let before = paragraph["spaceBeforePt"].as_f64().unwrap_or_default();
    let after = paragraph["spaceAfterPt"].as_f64().unwrap_or_default();
    assert!(
        (before - 14.0).abs() < 0.51 && (after - 14.0).abs() < 0.51,
        "自动段间距应取 14pt，实际 before={before} after={after}"
    );

    // 合成样本也要过"完备性"检查（不崩、部件齐全）
    for sample in [&chars_sample, &auto_sample] {
        let info = document_info(sample.to_string_lossy().to_string()).expect("合成样本应能解析");
        assert!(
            info.parts.iter().any(|part| part.name == "word/document.xml"),
            "合成样本应有 document.xml"
        );
        let xml = document_xml(sample.to_string_lossy().to_string()).expect("应能读 document.xml");
        assert!(xml.contains("<w:body"), "合成样本应有文档主体");
    }

    let _ = std::fs::remove_dir_all(&dir);
    eprintln!("合成样本校验通过：字符单位缩进 -> {indent}pt，自动段间距 -> {before}pt");
}