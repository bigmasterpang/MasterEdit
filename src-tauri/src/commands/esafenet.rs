//! Esafenet（亿赛通）透明加密文档的解密支持。
//!
//! 文件结构：4096 字节文件头 + 明文载荷（按 512 字节循环异或流加密）。
//! 头部偏移 12..16 为头长度（通常 0x1000），偏移 16..20 为标称原始大小。
//! 解密后若超出标称大小部分无法按 UTF-8 解析，说明是填充数据，需要截断。

/// 512 字节循环异或密钥流（十六进制，共 1024 个字符）
const KEY_HEX: &str = concat!(
    "d85a7b1e62d84117ec2d55103480eac3",
    "60c6369082c10ce2a7bf02f2bc9a93a9",
    "a998c7f7b1f560b7d53462570c5ee360",
    "2fe3ad2bdd754078ff8f46e332622fa8",
    "3d93f82f04b726e1fbedc01faf1dcf5b",
    "65e67de462229f2be1627439d8373586",
    "16f31ee841fb20db079ade21e9054199",
    "b86404da6bd797205e4073e2123e4285",
    "a1a3fada4c0c803de9de3e4d3f512148",
    "df9694c3635a0ddfa1413827190a8a2b",
    "3a276e539d755f1d4ba16fed462d72c1",
    "837b85cd2c6d28a708460010e8d6bdb7",
    "310d12cab0c7fc91bf4cbbae3b7ef94d",
    "f86c6ffa0561ba3282db3404198dccc4",
    "625830633b90d2a19f33c3ef8431cb95",
    "2dfe9bd69d6720a749b8e529c8268efc",
    "d6d73471cf15080447e2c41179686360",
    "7faa16e66c2f881d25e48c57f9fb8192",
    "c4ae70252a4865f55ea62fd0698f85f7",
    "f71c5d8c0347189ebe0dced25b19be30",
    "c40871abcc12315b7b3a92e83f09106b",
    "851d4d5beacb58503a45888429383c9b",
    "cb022c2eed2858891afdad68cd48088b",
    "ae24588d836d511dbf70104205ac34f6",
    "1e201e712958c8e5cccf3d8001cd11d2",
    "08332d22dbf189cbe1688bb46a380a0a",
    "198ca7750c52a66262101605853fec1c",
    "edc4d0ede6e0285ef0b88073efb89a0c",
    "7df476eab5e52b110d352d605f7f668d",
    "76988d26405f66a473f450ffae99a7cb",
    "19f958c04de5fb84ed6ea32edc7bac01",
    "c563b0b6b07457e5c10e2893e0da427e",
);

const KEY_LEN: usize = 512;

/// 解析十六进制密钥流（仅在首次解密时调用一次）。
fn key_bytes() -> [u8; KEY_LEN] {
    let mut key = [0u8; KEY_LEN];
    let hex = KEY_HEX.as_bytes();
    for (i, slot) in key.iter_mut().enumerate() {
        let hi = (hex[i * 2] as char).to_digit(16).unwrap_or(0) as u8;
        let lo = (hex[i * 2 + 1] as char).to_digit(16).unwrap_or(0) as u8;
        *slot = (hi << 4) | lo;
    }
    key
}

/// 判断是否为 Esafenet 加密文档。
///
/// 只认**结构**特征，绝不认「文件里出现过 Esafenet 字样」：
/// 旧实现用 `header.windows(8).any(|w| w == b"Esafenet")` 做兜底，导致任何在前 4096 字节里
/// 提到过该词的普通文本（例如本仓库自己的 `esafenet.rs`、`file.rs`）都会被误判成加密文档 ——
/// 打开时显示乱码，保存时按「加密文档」回写，真实内容被彻底破坏。
///
/// 真实加密文件的结构（已用样本 `负荷辨识加密.md`、`台区拓扑识别实验方案加密版.md` 验证）：
/// 前 4 字节固定为 `E0 A8 91 E7`，偏移 12..16 为头长度（0x1000），偏移 16..20 为标称明文长度。
pub fn is_esafenet_encrypted(data: &[u8]) -> bool {
    if data.len() < 4096 {
        return false;
    }
    if !data.starts_with(&[0xE0, 0xA8, 0x91, 0xE7]) {
        return false;
    }
    let header_size = u32::from_le_bytes([data[12], data[13], data[14], data[15]]) as usize;
    if header_size != 4096 && header_size > data.len() {
        return false;
    }
    let orig_size = u32::from_le_bytes([data[16], data[17], data[18], data[19]]) as usize;
    orig_size == 0 || orig_size <= data.len()
}

/// 解密 Esafenet 加密文档；不是加密文件时返回 None。
pub fn decrypt_esafenet(data: &[u8]) -> Option<Vec<u8>> {
    if !is_esafenet_encrypted(data) {
        return None;
    }
    let header_size = u32::from_le_bytes([data[12], data[13], data[14], data[15]]) as usize;
    let header_size = if header_size == 0 || header_size > data.len() {
        4096
    } else {
        header_size
    };
    let orig_size = u32::from_le_bytes([data[16], data[17], data[18], data[19]]) as usize;

    let key = key_bytes();
    let payload = &data[header_size..];
    let mut out = Vec::with_capacity(payload.len());
    for (i, byte) in payload.iter().enumerate() {
        out.push(byte ^ key[i % KEY_LEN]);
    }

    // 标称大小之后的填充字节无法按 UTF-8 解析时截断；是正文则完整保留
    if orig_size > 0 && orig_size < out.len() && std::str::from_utf8(&out).is_err() {
        out.truncate(orig_size);
    }
    Some(out)
}

/// 用 4096 字节文件头把明文重新加密（保存加密文档时使用）。
/// 头部沿用原文件的字节，仅把偏移 16..20 的标称原始大小更新为新明文长度。
pub fn encrypt_esafenet(header: &[u8], plaintext: &[u8]) -> Vec<u8> {
    let key = key_bytes();
    let mut head = vec![0u8; 4096];
    let copy_len = header.len().min(4096);
    head[..copy_len].copy_from_slice(&header[..copy_len]);
    head[16..20].copy_from_slice(&(plaintext.len() as u32).to_le_bytes());

    let mut out = Vec::with_capacity(4096 + plaintext.len());
    out.extend_from_slice(&head);
    for (i, byte) in plaintext.iter().enumerate() {
        out.push(byte ^ key[i % KEY_LEN]);
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    /// 用同一密钥流构造加密文件，验证解密可还原明文
    #[test]
    fn decrypts_roundtrip() {
        let plain = "# 标题\n\n正文内容 test\n".as_bytes();
        let key = key_bytes();
        let mut data = vec![0u8; 4096];
        data[0..4].copy_from_slice(&[0xE0, 0xA8, 0x91, 0xE7]);
        data[12..16].copy_from_slice(&4096u32.to_le_bytes());
        data[16..20].copy_from_slice(&(plain.len() as u32).to_le_bytes());
        for (i, byte) in plain.iter().enumerate() {
            data.push(byte ^ key[i % KEY_LEN]);
        }
        let out = decrypt_esafenet(&data).expect("应识别为加密文档");
        assert_eq!(out, plain);
    }

    /// 普通 Markdown 不应被误判
    #[test]
    fn ignores_plain_files() {
        assert!(decrypt_esafenet(b"# hello\nworld\n").is_none());
        assert!(!is_esafenet_encrypted(b"short"));
    }

    /// 回归：前 4096 字节里出现 "Esafenet" 字样的普通长文本不得被误判。
    /// 旧实现用 `windows(8).any(|w| w == b"Esafenet")` 兜底，导致本仓库的 esafenet.rs / file.rs
    /// 这类源码在应用里被当成加密文档打开（乱码），保存时更会把正文彻底改写。
    #[test]
    fn ignores_plain_text_mentioning_esafenet() {
        let mut data = b"//! Esafenet transparent encryption helper\n".to_vec();
        data.resize(9000, b'a');
        assert!(!is_esafenet_encrypted(&data));
        assert!(decrypt_esafenet(&data).is_none());
    }

    /// 头部特征字节正确但结构字段不成立时也不认（防止极端巧合）
    #[test]
    fn rejects_broken_header_structure() {
        let mut data = vec![0u8; 9000];
        data[0..4].copy_from_slice(&[0xE0, 0xA8, 0x91, 0xE7]);
        data[12..16].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // 头长度超出文件
        data[16..20].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // 标称明文长度超出文件
        assert!(!is_esafenet_encrypted(&data));
    }

    /// 重新加密后可再次解密（保存加密文档的核心保障）
    #[test]
    fn encrypt_then_decrypt_roundtrip() {
        let key = key_bytes();
        let mut header = vec![0u8; 4096];
        header[0..4].copy_from_slice(&[0xE0, 0xA8, 0x91, 0xE7]);
        header[8..12].copy_from_slice(&512u32.to_le_bytes());
        header[12..16].copy_from_slice(&4096u32.to_le_bytes());

        let plain = "# 标题\n\n保存后仍是加密格式。\n".repeat(3);
        let encrypted = encrypt_esafenet(&header, plain.as_bytes());
        // 文件头原样保留，标称大小已更新
        assert_eq!(&encrypted[..4], &[0xE0, 0xA8, 0x91, 0xE7]);
        let size_field = u32::from_le_bytes([encrypted[16], encrypted[17], encrypted[18], encrypted[19]]);
        assert_eq!(size_field as usize, plain.len());
        // 二次解密还原明文
        let back = decrypt_esafenet(&encrypted).expect("应能再次解密");
        assert_eq!(back, plain.as_bytes());
        let _ = key;
    }

    /// 用真实加密文件做一次「解密 → 改内容 → 重新加密 → 落盘 → 再读回解密」验证（文件不存在时跳过）
    #[test]
    fn real_file_roundtrip_if_present() {
        let path = r"C:\opencode\test\负荷辨识加密.md";
        let Ok(raw) = std::fs::read(path) else {
            return;
        };
        let plain = decrypt_esafenet(&raw).expect("真实文件应能解密");
        let modified = format!("{}\n\n<!-- 追加测试 -->\n", String::from_utf8_lossy(&plain));
        let encrypted = encrypt_esafenet(&raw[..4096], modified.as_bytes());

        // 走一次真实文件系统写入，确认落盘内容仍可解密
        let temp = std::env::temp_dir().join("masteredit-esafenet-roundtrip.md");
        std::fs::write(&temp, &encrypted).expect("写入临时文件");
        let read_back = std::fs::read(&temp).expect("读回临时文件");
        let back = decrypt_esafenet(&read_back).expect("落盘后应仍能解密");
        assert_eq!(String::from_utf8_lossy(&back), modified);
        let _ = std::fs::remove_file(&temp);
    }
}
