# -*- coding: utf-8 -*-
"""
020 Email download.py
功能：
1) 读取 .env 中 QQ 邮箱 IMAP 凭据，登录并抓取最近 N 封邮件。
2) 以配置的关键词筛选两类邮件：
   - KEYWORDS["waiting"]      → “等待您查看”
   - KEYWORDS["heyu_da"]      → “合肥市和裕达”
   各自选取【最新】一封。
3) 提取选中邮件的 HTML 正文（用于后续解析表格），并：
   - 若命中 heyu_da 类，下载其附件到保存目录（文件名追加时间戳，保留原扩展名）。
4) 将“选中的 subject + 收到时间（ISO8601）”写入保存目录下的 mail_meta.json，供后续脚本读取。
5) 解析 HTML 中首个合理表格并导出 Excel（第一页）。
使用：
- 可传入保存目录作为第1个命令行参数；不传则使用平台默认目录（Windows: ./data；其他: ~/data）。
"""

import os
import sys
import re
import time
import platform
import json
import email
import imaplib
from email.header import decode_header
from email.utils import parsedate_tz, mktime_tz
from datetime import datetime

import pandas as pd
import openpyxl
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment
from dotenv import load_dotenv


# ================================
# 📂 路径配置（支持主程序传参）
# ================================
if platform.system() == "Windows":
    default_save_path = os.path.join(os.getcwd(), "data")  # Windows: 相对路径 ./data
else:
    default_save_path = os.path.expanduser("~/data")       # Linux/macOS: 家目录 ~/data

excel_save_path = sys.argv[1] if len(sys.argv) >= 2 else default_save_path
os.makedirs(excel_save_path, exist_ok=True)
print(f"📂 保存路径: {os.path.abspath(excel_save_path)}")

# ================================
# 🔧 关键词/邮箱配置（集中管理）
# ================================
KEYWORDS = {
    "waiting": "等待您查看",
    "heyu_da": "合肥市和裕达",
}
MAILBOX = os.getenv("IMAP_MAILBOX", "INBOX")          # QQ/多数 IMAP 兼容 "INBOX"
RECENT_LIMIT = int(os.getenv("RECENT_LIMIT", "15"))  # 最近抓取邮件数量上限
META_FILENAME = "mail_meta.json"                      # 元数据文件名（写入 excel_save_path）

# ================================
# 📧 邮箱凭据（.env）
# ================================
load_dotenv()  # 可改为 load_dotenv(dotenv_path="...") 定点加载

email_user = os.getenv("EMAIL_ADDRESS_QQ")
# 兼容旧写法 EMAIL_PASSWOR_QQ（少了D）
email_password = os.getenv("EMAIL_PASSWORD_QQ") or os.getenv("EMAIL_PASSWOR_QQ")
email_server = os.getenv("IMAP_SERVER", "imap.qq.com")

if not email_user or not email_password:
    raise ValueError("❌ 环境变量未正确配置（EMAIL_ADDRESS_QQ / EMAIL_PASSWORD_QQ）！")

print("📬 正在使用邮箱:", email_user)

# ================================
# 🔑 标题解码与清理
# ================================
def decode_str(s: str) -> str:
    if not s:
        return ""
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    elif isinstance(value, bytes):
        value = value.decode("utf-8", errors="ignore")
    return value

def clean_subject(subject: str) -> str:
    cleaned_subject = re.sub(r'\[([^\[\]]+)\]', r'\1', subject or "")
    cleaned_subject = re.sub(r'【([^【】]+)】', r'\1', cleaned_subject)
    return cleaned_subject.strip()

# ================================
# 📨 抓取邮件并输出 HTML/元数据/附件
# ================================
def fetch_html_from_emails(server: str, user: str, password: str, save_dir: str) -> str | None:
    mail = None
    html_content = None

    # 预置元数据（两类最新邮件）
    meta = {
        "selected_heyu_da_subject": None,
        "selected_heyu_da_received_at": None,
        "selected_waiting_subject": None,
        "selected_waiting_received_at": None,
    }

    try:
        print("🔗 正在连接邮箱...")
        mail = imaplib.IMAP4_SSL(server)
        mail.login(user, password)

        # 选择邮箱目录
        status, _ = mail.select(MAILBOX)
        if status != "OK":
            print(f"⚠️ 无法选择邮箱目录 {MAILBOX}，尝试使用 INBOX")
            mail.select("INBOX")

        print(f"🔎 正在检索最近 {RECENT_LIMIT} 封邮件...")
        status, messages = mail.search(None, "ALL")
        if status != "OK":
            print("未找到邮件")
            return None

        mail_ids = messages[0].split()
        if not mail_ids:
            print("邮箱为空。")
            return None

        recent_mail_ids = mail_ids[-RECENT_LIMIT:]
        print(f"📨 共 {len(mail_ids)} 封，处理最近 {len(recent_mail_ids)} 封。")

        inventory_query_emails = []

        # 遍历最近的 N 封邮件
        for i, mail_id in enumerate(recent_mail_ids, start=1):
            status, msg_data = mail.fetch(mail_id, "(RFC822)")
            if status != "OK" or not msg_data or not msg_data[0]:
                print(f"⚠️ 第 {i} 封抓取失败")
                continue

            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = decode_str(msg.get("Subject"))
            from_ = decode_str(msg.get("From"))
            date_raw = decode_str(msg.get("Date"))

            # 转换为 datetime，失败则兜底为 1970-01-01
            mail_date = parsedate_tz(date_raw)
            if mail_date:
                mail_datetime = datetime.fromtimestamp(mktime_tz(mail_date))
            else:
                mail_datetime = datetime(1970, 1, 1)

            cleaned_subject = clean_subject(subject)

            print(f"  · 第 {i} 封 | 原: {subject} | 清理: {cleaned_subject} | 发件人: {from_} | 收到: {mail_datetime}")

            # 仅收集标题命中两类关键词的邮件
            if (KEYWORDS["waiting"] in cleaned_subject) or (KEYWORDS["heyu_da"] in cleaned_subject):
                inventory_query_emails.append({
                    "mail_id": mail_id,
                    "subject": subject,
                    "cleaned_subject": cleaned_subject,
                    "date": mail_datetime,
                    "msg": msg
                })

        # 打印筛选列表
        if inventory_query_emails:
            print("\n✅ 命中关键词的邮件：")
            for item in inventory_query_emails:
                print(f"  - {item['cleaned_subject']} | {item['date']}")
        else:
            print("\nℹ️ 未命中任何关键词邮件。")

        # 选出“合肥市和裕达”最新一封 → 提取 HTML + 下载附件
        selected_heyu = _pick_latest(inventory_query_emails, KEYWORDS["heyu_da"])
        if selected_heyu:
            html_content = extract_html_from_msg(selected_heyu["msg"]) or html_content
            print(f"\n📌 选中(合肥市和裕达): {selected_heyu['cleaned_subject']} | {selected_heyu['date']}")
            meta["selected_heyu_da_subject"] = selected_heyu["cleaned_subject"]
            meta["selected_heyu_da_received_at"] = selected_heyu["date"].isoformat() if selected_heyu["date"] else None
            # 下载附件
            download_attachments(selected_heyu["msg"], save_dir)

        # 选出“等待您查看”最新一封 → 提取 HTML
        selected_waiting = _pick_latest(inventory_query_emails, KEYWORDS["waiting"])
        if selected_waiting:
            html_content = extract_html_from_msg(selected_waiting["msg"]) or html_content
            print(f"\n📌 选中(等待您查看): {selected_waiting['cleaned_subject']} | {selected_waiting['date']}")
            meta["selected_waiting_subject"] = selected_waiting["cleaned_subject"]
            meta["selected_waiting_received_at"] = selected_waiting["date"].isoformat() if selected_waiting["date"] else None

        # 写出元数据
        _write_meta(meta, os.path.join(save_dir, META_FILENAME))

        if html_content:
            print("✅ 已获取选定邮件的 HTML 正文。")
        else:
            print("ℹ️ 未找到符合条件的 HTML 正文。")

        return html_content

    except imaplib.IMAP4.error as e:
        print(f"IMAP 错误: {e}")
        return None
    except Exception as e:
        print(f"获取邮件失败: {e}")
        return None
    finally:
        try:
            if mail is not None:
                mail.logout()
        except Exception:
            pass

def _pick_latest(candidates: list[dict], keyword: str) -> dict | None:
    """在 candidates 中选出 cleaned_subject 含 keyword 的【最新】一封。"""
    selected = None
    for item in candidates:
        if keyword in item["cleaned_subject"]:
            if (selected is None) or (item["date"] > selected["date"]):
                selected = item
    return selected

# ================================
# 🧩 从邮件中提取 HTML 正文
# ================================
def extract_html_from_msg(msg) -> str | None:
    html_content = None
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition") or "")
            if content_type == "text/html" and "attachment" not in content_disposition:
                charset = part.get_content_charset() or part.get_charset() or "utf-8"
                try:
                    html_content = part.get_payload(decode=True).decode(charset, errors="ignore")
                except Exception:
                    html_content = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                break
    else:
        if msg.get_content_type() == "text/html":
            charset = msg.get_content_charset() or msg.get_charset() or "utf-8"
            try:
                html_content = msg.get_payload(decode=True).decode(charset, errors="ignore")
            except Exception:
                html_content = msg.get_payload(decode=True).decode("utf-8", errors="ignore")
    return html_content

# ================================
# 📎 下载附件（追加时间戳，保留扩展名）
# ================================
def download_attachments(msg, download_folder: str) -> None:
    """下载邮件附件：文件名按原名+时间戳，保留扩展名；若无扩展名则根据 MIME 猜测。"""
    if not msg.is_multipart():
        return

    import mimetypes
    import unicodedata
    from email.header import decode_header

    def _decode_filename(raw: str) -> str:
        """将可能被拆分编码的文件名各段解码并拼接；规范化全角点等。"""
        parts = decode_header(raw)
        s = ""
        for p, enc in parts:
            if isinstance(p, bytes):
                s += p.decode(enc or "utf-8", errors="ignore")
            else:
                s += p
        s = unicodedata.normalize("NFC", s).replace("．", ".").strip().strip(".")
        return s

    def _sanitize(name: str) -> str:
        """清理不合法文件名字符。"""
        invalid = '<>:"/\\|?*'
        name = "".join((c if c not in invalid else "_") for c in name)
        # 避免隐藏名或空名
        name = name.strip().strip(".")
        return name or "attachment"

    def _guess_ext(content_type: str) -> str:
        """根据 MIME 猜测扩展名，内置常见兜底。"""
        overrides = {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
            "application/vnd.ms-excel": ".xls",
            "text/csv": ".csv",
            "application/zip": ".zip",
            "application/pdf": ".pdf",
        }
        return overrides.get(content_type) or (mimetypes.guess_extension(content_type) or "")

    def _ensure_unique(path: str) -> str:
        """如重名则在基名末尾追加(_2/_3...)避免覆盖。"""
        if not os.path.exists(path):
            return path
        base, ext = os.path.splitext(path)
        i = 2
        while True:
            candidate = f"{base}_{i}{ext}"
            if not os.path.exists(candidate):
                return candidate
            i += 1

    for part in msg.walk():
        # 跳过容器部件，仅处理真正内容/附件
        if part.get_content_maintype() == "multipart":
            continue

        content_disposition = str(part.get("Content-Disposition") or "")
        raw_name = part.get_filename()  # Python 会处理 RFC2231 的 filename* 情况

        # 既不是附件也没有文件名的，跳过
        if "attachment" not in content_disposition and not raw_name:
            continue

        # 1) 解析文件名
        if raw_name:
            filename = _decode_filename(raw_name)
        else:
            # 没有文件名，用类型生成占位名
            filename = f"attachment{_guess_ext(part.get_content_type())}"

        # 2) 拆分扩展名；若缺失则根据 MIME 猜测
        base_name, ext = os.path.splitext(filename)
        if not ext:
            ext = _guess_ext(part.get_content_type())

        # 3) 追加时间戳并清理文件名
        ts = time.strftime("%Y%m%d_%H%M%S")
        safe_base = _sanitize(base_name)
        safe_name = f"{safe_base}_{ts}{ext}"
        file_path = os.path.join(download_folder, safe_name)
        file_path = _ensure_unique(file_path)

        # 4) 写入磁盘
        file_data = part.get_payload(decode=True)
        if not file_data:
            continue
        with open(file_path, "wb") as f:
            f.write(file_data)

        print(f"📥 附件已下载: {file_path}")

# ================================
# 🧠 解析 HTML 表格并导出 Excel
# ================================
# ================================
# 🧠 解析 HTML 表格（改为仅 BeautifulSoup）
# ================================
def parse_html_table(html_content: str) -> list[list[str]]:
    print("正在解析 HTML 内容中的表格...")

    # 可选：保存快照便于排查 GitHub Actions 上的解析结果
    try:
        snap_path = os.path.join(excel_save_path, "last_mail_html.html")
        with open(snap_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        print(f"🔎 HTML 快照: {snap_path}")
    except Exception:
        pass

    soup = BeautifulSoup(html_content, "html.parser")
    table = soup.find("table")
    if not table:
        print("未找到 HTML 表格！")
        return []

    data = []
    rows = table.find_all("tr")
    header = None

    for idx, row in enumerate(rows):
        cols = [ele.get_text(strip=True) for ele in row.find_all(["td", "th"])]

        if not cols:
            print(f"第 {idx + 1} 行是空行，跳过")
            continue

        # 首行特判：列数过多当成正文，跳过
        if header is None:
            if idx == 0 and len(cols) > 10:
                print("第一行列数过多，认为其为正文内容，跳过")
                continue
            header = cols
            data.append(header)
            continue

        if len(cols) != len(header):
            print(f"第 {idx + 1} 行列数与表头不匹配，跳过")
            continue
        if cols == header:
            print(f"第 {idx + 1} 行是重复表头，跳过")
            continue

        data.append(cols)

    print(f"成功提取 {len(data)} 行表格数据。")

    # 保留前导零：纯数字以字符串写入
    for i in range(len(data)):
        for j in range(len(data[i])):
            if isinstance(data[i][j], str) and data[i][j].isdigit():
                data[i][j] = str(data[i][j])

    return data



def save_to_excel(data: list[list[str]], save_dir: str, file_prefix="存量查询") -> None:
    if not data:
        print("ℹ️ 没有可导出的数据。")
        return

    # 去重（保序）
    seen = set()
    unique_data = []
    for row in data:
        tup = tuple(row)
        if tup not in seen:
            seen.add(tup)
            unique_data.append(row)

    df = pd.DataFrame(unique_data)

    # 文件名加时间戳
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    file_name = f"{file_prefix}_{timestamp}.xlsx"
    full_path = os.path.join(save_dir, file_name)

    print(f"💾 正在保存 Excel：{full_path}")
    # 不写 header（原逻辑）
    df.to_excel(full_path, index=False, header=False)

    # openpyxl 再细化格式
    wb = openpyxl.load_workbook(full_path)
    ws = wb.active
    ws.title = "第一页"

    # 示例：尝试对第5列、第6列应用千分位 & 右对齐（仅当多数可被识别为数值）
    decimal_columns = [4, 5]  # 0-based 索引，对应第5/6列
    max_row = ws.max_row
    for col in decimal_columns:
        # 统计数值比例
        numeric_count = 0
        for r in range(2, max_row + 1):
            val = ws.cell(row=r, column=col + 1).value
            try:
                float(str(val).replace(",", ""))  # 尝试可转数
                numeric_count += 1
            except Exception:
                pass
        # 超过一半可视为数值 → 应用格式
        if numeric_count >= (max_row - 1) / 2:
            for r in range(2, max_row + 1):
                cell = ws.cell(row=r, column=col + 1)
                try:
                    v = float(str(cell.value).replace(",", ""))
                    cell.value = v
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right')
                except Exception:
                    # 无法转数值就跳过
                    pass

    wb.save(full_path)
    print("✅ Excel 保存完成。")

def _write_meta(meta: dict, path: str) -> None:
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        print(f"📝 元数据已写入: {path}")
    except Exception as e:
        print(f"⚠️ 元数据写入失败: {e}")

# ================================
# 🚀 主程序
# ================================
if __name__ == '__main__':
    print("程序启动")
    html_content = fetch_html_from_emails(email_server, email_user, email_password, excel_save_path)

    if html_content:
        preview = html_content[:400].replace("\n", " ")
        print(f"HTML 预览: {preview} ...")

        table_data = parse_html_table(html_content)
        if table_data:
            save_to_excel(table_data, excel_save_path, file_prefix="存量查询")
        else:
            print("表格为空，未导出 Excel。")
    else:
        print("未获取到 HTML，程序结束。")
