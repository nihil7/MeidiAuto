import os
import time
import random
import cloudscraper
from datetime import datetime
from dotenv import load_dotenv

# === 加载环境变量（仅限本地）===
if not os.getenv("GITHUB_ACTIONS", "").lower() == "true":
    load_dotenv()

# === 请求头信息 ===
headers = {
    'User-Agent': 'Mozilla/5.0',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Connection': 'keep-alive'
}

# === Cloudflare 兼容请求对象 ===
scraper = cloudscraper.create_scraper()

# === 统一配置所有网站 ===
SITES = {
    "PTTIME": {
        "url": "https://www.pttime.org/attendance.php?type=sign&uid={uid}",
        "accounts": {
            "A": {"uid": "2785"},
            "B": {"uid": "20801"}
        }
    },
    "XXXFORUM": {
    "url": "https://1ptba.com/attendance.php",  # 没有 {uid}
    "accounts": {
        "X1": {},
        "X2": {}
    }
}
}

# === 读取某账号的 Cookie ===
def build_cookie(site, account):
    prefix = f"{site}_{account}".upper()
    return {
        'logged_in': os.getenv(f'{prefix}_LOGGED_IN'),
        'cf_clearance': os.getenv(f'{prefix}_CF_CLEARANCE'),
        'c_secure_uid': os.getenv(f'{prefix}_C_SECURE_UID'),
        'c_secure_tracker_ssl': os.getenv(f'{prefix}_C_SECURE_TRACKER_SSL'),
        'c_secure_ssl': os.getenv(f'{prefix}_C_SECURE_SSL'),
        'c_secure_pass': os.getenv(f'{prefix}_C_SECURE_PASS'),
        'c_secure_login': os.getenv(f'{prefix}_C_SECURE_LOGIN'),
        'c_lang_folder': os.getenv(f'{prefix}_C_LANG_FOLDER'),
    }

# === Cookie 检查 ===
def is_cookie_valid(cookie_dict):
    return all(cookie_dict.values())

# === 打印 Cookie 简要 ===
def print_cookie_info(name, cookie_dict):
    print(f"🍪 {name} cookies:")
    for k, v in cookie_dict.items():
        print(f"  {k} = {v if v else '❌ 缺失'}")

# === 执行签到 ===
def check_in(full_url, cookies, site, account):
    print(f"🚀 正在签到 [{site} - {account}]: {full_url}")
    try:
        response = scraper.get(full_url, cookies=cookies, headers=headers, timeout=10)
        print(f"✅ 状态码: {response.status_code}")
        print(f"📄 内容（前200字）: {response.text[:200]}...")
        if response.status_code == 200:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"🎉 成功签到：{site} - {account} @ {now}")
        else:
            print(f"❌ 签到失败：{site} - {account}")
    except Exception as e:
        print(f"🚨 请求异常：{site} - {account} - {str(e)}")

# === 主函数 ===
def main():
    for site, site_info in SITES.items():
        url_template = site_info["url"]
        for account, info in site_info["accounts"].items():
            cookie = build_cookie(site, account)
            print_cookie_info(f"{site}-{account}", cookie)

            if not is_cookie_valid(cookie):
                print(f"⚠️ 跳过：{site} - {account}，Cookie 信息不完整")
                continue

            full_url = url_template.format(**info) if '{' in url_template else url_template
            check_in(full_url, cookie, site, account)

            delay = random.uniform(0, 5)
            print(f"⏳ 等待 {delay:.2f} 秒...\n")
            time.sleep(delay)

if __name__ == "__main__":
    main()
