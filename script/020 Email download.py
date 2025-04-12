import os
import platform
import re
import sys
import time
from email.header import decode_header

import openpyxl
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment
from dotenv import load_dotenv

# ================================
# 📂 路径配置（支持主程序传参）
# ================================
if platform.system() == "Windows":
    default_save_path = os.path.join(os.getcwd(), "data",)  # Windows 用相对路径
else:
    default_save_path = os.path.expanduser("~/date")  # Linux/macOS

# 允许主程序传参
excel_save_path = sys.argv[1] if len(sys.argv) >= 2 else default_save_path
print(f"📂 保存路径: {excel_save_path}")

# 确保路径存在
os.makedirs(excel_save_path, exist_ok=True)




# ================================
# 邮箱配置（保密信息建议放配置文件）
# ================================
# 加载 .env 文件中的变量
load_dotenv()

# 从环境变量中读取邮箱和授权码
email_user = os.getenv("EMAIL_ADDRESS_QQ")
email_password = os.getenv("EMAIL_PASSWOR_QQ")  # 注意变量名拼写！

if not email_user or not email_password:
    raise ValueError("❌ 环境变量未正确配置，无法获取邮箱账户或密码！")

# 以下是你原本的逻辑
print("📬 正在使用邮箱:", email_user)

email_server = 'imap.qq.com'


# ================================
# 🔑 邮件标题解码
# ================================
def decode_str(s):
    if not s:
        return ''
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    elif isinstance(value, bytes):
        value = value.decode('utf-8')
    return value

# 邮件标题清理函数
def clean_subject(subject):
    # 替换 [xxx] 为 xxx（保留内容，删除括号）
    cleaned_subject = re.sub(r'\[([^\[\]]+)\]', r'\1', subject)
    # 替换 【xxx】为 xxx（保留内容，删除括号）
    cleaned_subject = re.sub(r'【([^\【】]+)】', r'\1', cleaned_subject)
    # 去除前后空格
    cleaned_subject = cleaned_subject.strip()
    return cleaned_subject

import email
import imaplib
from email.utils import parsedate_tz, mktime_tz
from datetime import datetime

def fetch_html_from_emails(server, user, password):
    try:
        print("正在连接邮箱...")
        mail = imaplib.IMAP4_SSL(server)
        mail.login(user, password)

        mail.select('inbox')

        print("正在检索最近6封邮件...")
        status, messages = mail.search(None, 'ALL')
        if status != 'OK':
            print("未找到邮件")
            return None

        mail_ids = messages[0].split()
        recent_mail_ids = mail_ids[-15:]

        print(f"共找到 {len(mail_ids)} 封邮件，正在处理最近6封邮件。")

        html_content = None
        inventory_query_emails = []

        # 遍历最近的6封邮件
        for i, mail_id in enumerate(recent_mail_ids):
            status, msg_data = mail.fetch(mail_id, '(RFC822)')
            if status != 'OK':
                print(f"邮件 {i + 1} 抓取失败")
                continue

            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = decode_str(msg.get('Subject'))
            from_ = decode_str(msg.get('From'))
            date_ = decode_str(msg.get('Date'))

            # 将邮件日期转换为 datetime 对象
            mail_date = parsedate_tz(date_)
            if mail_date:
                mail_datetime = datetime.fromtimestamp(mktime_tz(mail_date))

            # 清理邮件主题
            cleaned_subject = clean_subject(subject)

            print(f"第 {i + 1} 封邮件 - 原主题: {subject} | 清理后主题: {cleaned_subject} | 发件人: {from_} | 日期: {mail_datetime}")

            # 筛选标题包含“等待您查看”和“合肥市和裕达”的邮件
            if "等待您查看" in cleaned_subject or "合肥市和裕达" in cleaned_subject:
                # 满足条件的邮件，添加到筛选列表
                inventory_query_emails.append({
                    'mail_id': mail_id,
                    'subject': subject,
                    'cleaned_subject': cleaned_subject,
                    'date': mail_datetime,
                    'msg': msg
                })

        # 打印筛选的邮件列表
        print("\n筛选出的邮件:")
        for email_data in inventory_query_emails:
            print(f"主题: {email_data['cleaned_subject']} | 日期: {email_data['date']}")

        # 从筛选出的邮件中，找出最新包含“合肥市和裕达”的邮件
        selected_email = None
        for email_data in inventory_query_emails:
            if "合肥市和裕达" in email_data['cleaned_subject']:
                if not selected_email or email_data['date'] > selected_email['date']:
                    selected_email = email_data

        if selected_email:
            html_content = extract_html_from_msg(selected_email['msg'])
            print(f"\n选中邮件: {selected_email['cleaned_subject']} | 日期: {selected_email['date']}")

            # 下载附件
            download_attachments(selected_email['msg'], excel_save_path)

        # 从筛选出的邮件中，找出最新包含“等待您查看”的邮件
        selected_email = None
        for email_data in inventory_query_emails:
            if "等待您查看" in email_data['cleaned_subject']:
                if not selected_email or email_data['date'] > selected_email['date']:
                    selected_email = email_data

        if selected_email:
            html_content = extract_html_from_msg(selected_email['msg'])
            print(f"\n选中邮件: {selected_email['cleaned_subject']} | 日期: {selected_email['date']}")

        mail.logout()

        if html_content:
            print("成功获取选定邮件的 HTML 内容！")
        else:
            print("未找到符合条件的邮件 HTML 内容。")

        return html_content

    except imaplib.IMAP4.error as e:
        print(f"IMAP 错误: {e}")
        return None
    except Exception as e:
        print(f"获取邮件失败: {e}")
        return None


# 提取 HTML 正文
def extract_html_from_msg(msg):
    html_content = None

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition") or "")

            if content_type == "text/html" and "attachment" not in content_disposition:
                charset = part.get_content_charset() or part.get_charset() or 'utf-8'
                try:
                    html_content = part.get_payload(decode=True).decode(charset, errors='ignore')
                except Exception as e:
                    print(f"HTML 解码失败，尝试 utf-8：{e}")
                    html_content = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                break
    else:
        content_type = msg.get_content_type()
        print(f"[单部分邮件] 当前内容类型: {content_type}")

        if content_type == "text/html":
            print("判断结果：这是 HTML 内容，继续处理！")
            charset = msg.get_content_charset() or msg.get_charset() or 'utf-8'
            try:
                html_content = msg.get_payload(decode=True).decode(charset, errors='ignore')
            except Exception as e:
                print(f"HTML 解码失败，尝试 utf-8：{e}")
                html_content = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
        else:
            print("判断结果：这不是 HTML 内容，跳过处理。")

    return html_content





# 下载邮件附件并加入日期时间
def download_attachments(msg, download_folder):
    if msg.is_multipart():
        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition"))
            if "attachment" in content_disposition:
                # 解码附件文件名
                filename, encoding = decode_header(part.get_filename())[0]

                # 如果文件名是字节类型，则转换为字符串
                if isinstance(filename, bytes):
                    filename = filename.decode(encoding if encoding else 'utf-8')

                # 获取当前日期和时间
                current_time = time.strftime("%Y%m%d_%H%M%S")

                # 确保附件保存为 .xlsx 格式，并在文件名中加入日期和时间
                base_name, ext = os.path.splitext(filename)
                if not filename.endswith('.xlsx'):
                    filename = f"{base_name}_{current_time}.xlsx"
                else:
                    filename = f"{base_name}_{current_time}{ext}"

                # 打印附件的完整名称
                print(f"附件名: {filename}")

                # 下载附件
                file_data = part.get_payload(decode=True)
                file_path = os.path.join(download_folder, filename)

                # 保存附件
                with open(file_path, "wb") as f:
                    f.write(file_data)

                # 打印附件的完整路径
                print(f"附件已下载到: {file_path}")


# 解析 HTML 表格内容
def parse_html_table(html_content):
    print("正在解析 HTML 内容中的表格...")
    soup = BeautifulSoup(html_content, "html.parser")
    table = soup.find("table")

    if not table:
        print("未找到 HTML 表格！")
        return []

    data = []
    rows = table.find_all("tr")

    header = None
    for idx, row in enumerate(rows):
        cols = row.find_all(["td", "th"])
        cols = [ele.get_text(strip=True) for ele in cols]

        # 只打印特殊情况下的行，其他行不打印
        if not cols:
            print(f"第 {idx + 1} 行是空行，跳过")
            continue

        # 处理第一行
        if idx == 0:
            if len(cols) > 10:
                # 第一行列数过多，认为其为正文内容，跳过
                print("第一行列数过多，认为其为正文内容，跳过")
                continue  # 跳过第一行
            else:
                header = cols
                data.append(header)  # 添加表头
                continue  # 不打印第一行

        elif header is None:
            # 还未找到表头
            header = cols
            data.append(header)
            continue  # 不打印表头的副本

        else:
            if len(cols) != len(header):
                # 列数与表头不匹配，跳过该行
                print(f"第 {idx + 1} 行列数与表头不匹配，跳过")
                continue

            if cols == header:
                # 重复表头，跳过
                print(f"第 {idx + 1} 行是重复表头，跳过")
                continue

            # 添加符合条件的行
            data.append(cols)

    print(f"成功提取 {len(data)} 行表格数据。")

    # 处理数据：将包含前导零的字段转为字符串，以确保前导零不丢失
    for i in range(len(data)):
        for j in range(len(data[i])):
            # 检查是否是需要保留前导零的字段
            if isinstance(data[i][j], str) and data[i][j].isdigit():
                data[i][j] = str(data[i][j])

    return data

# 保存为 Excel 文件并格式化
def save_to_excel(data, file_prefix="存量查询"):
    if not data:
        print("没有数据可以保存，跳过 Excel 导出。")
        return

    print("正在处理表格数据（去重）...")
    # 去重（保持顺序）
    unique_data = []
    seen = set()

    for row in data:
        row_tuple = tuple(row)
        if row_tuple not in seen:
            seen.add(row_tuple)
            unique_data.append(row)

    # 创建 Excel 文件
    df = pd.DataFrame(unique_data)

    # 文件名加时间戳
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    file_name = f"{file_prefix}_{timestamp}.xlsx"
    full_path = os.path.join(excel_save_path, file_name)

    # 将包含前导零的字段转为字符串，以确保前导零不丢失
    for col in df.columns:
        df[col] = df[col].apply(lambda x: str(x) if isinstance(x, int) or isinstance(x, float) else x)

    # 保存 Excel 文件
    print(f"正在保存 Excel 文件：{file_name} ...")
    df.to_excel(full_path, index=False, header=False)

    # 使用 openpyxl 进一步格式化 Excel 文件
    wb = openpyxl.load_workbook(full_path)
    ws = wb.active
    ws.title = "第一页"  # 设置工作表名称

    # 检查第5列和第6列是否需要格式化为千分位
    decimal_columns = [4, 5]  # 第5列和第6列对应索引为4和5

    for col in decimal_columns:
        column_data = [str(ws.cell(row=row_num, column=col + 1).value) for row_num in range(2, len(data) + 1)]
        count_with_decimal = sum(1 for value in column_data if '.' in value)
        if count_with_decimal > len(column_data) / 2:  # 如果超过一半含有小数点
            for row_num in range(2, len(data) + 1):
                cell = ws.cell(row=row_num, column=col + 1)
                cell.number_format = '#,##0.00'  # 千分位格式
                cell.alignment = Alignment(horizontal='right')  # 右对齐

    # 保存格式化后的 Excel 文件
    wb.save(full_path)


# 主程序
if __name__ == '__main__':
    print("程序启动！")

    html_content = fetch_html_from_emails(email_server, email_user, email_password)

    if html_content:
        print("\nHTML 内容部分预览：")
        print(html_content[:500])  # 打印前500字符预览

        table_data = parse_html_table(html_content)

        if table_data:
            save_to_excel(table_data, file_prefix="存量查询")
        else:
            print("表格数据为空，未导出 Excel。")
    else:
        print("未获取到 HTML 内容，程序结束。")
