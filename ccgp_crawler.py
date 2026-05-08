# -*- coding: utf-8 -*-
"""
中国政府采购网招标公告爬虫
使用 requests + BeautifulSoup
"""

import csv
import re
import time
import urllib.parse
import requests
from bs4 import BeautifulSoup


# 请求头，模拟浏览器访问
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Referer": "https://search.ccgp.gov.cn/",
    "Connection": "keep-alive",
}

# 搜索关键词
KEYWORD = "医疗器械"

# 列表页 URL 模板
LIST_URL_TEMPLATE = (
    "https://search.ccgp.gov.cn/bxsearch?"
    "searchtype=1"
    "&page_index={page}"
    "&start_time="
    "&end_time="
    "&timeType=2"
    "&searchparam="
    "&searchchannel=0"
    "&dbselect=bidx"
    "&kw={kw}"
    "&bidSort=0"
    "&pinMu=0"
    "&bidType=0"
    "&buyerName="
    "&projectId="
    "&displayZone="
    "&zoneId="
    "&agentName="
)


def fetch_html(url, retries=3):
    """获取页面 HTML，失败时自动重试"""
    for attempt in range(1, retries + 1):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=30, allow_redirects=True)
            resp.encoding = resp.apparent_encoding or "utf-8"
            if resp.status_code == 200:
                # 检查是否被反爬拦截
                if "频繁访问" in resp.text or "您的访问过于频繁" in resp.text:
                    print(f"  [!] 访问过于频繁，等待 5 秒后重试 ({attempt}/{retries})")
                    time.sleep(5)
                    continue
                return resp.text
            else:
                print(f"  [!] HTTP {resp.status_code}，重试 ({attempt}/{retries})")
        except Exception as e:
            print(f"  [!] 请求异常: {e}，重试 ({attempt}/{retries})")
        time.sleep(2)
    return None


def parse_list_page(html):
    """解析列表页，返回 [(标题, 发布时间, 详情页链接), ...]"""
    soup = BeautifulSoup(html, "html.parser")
    ul = soup.find("ul", class_="vT-srch-result-list-bid")
    if not ul:
        return []

    items = []
    for li in ul.find_all("li"):
        a_tag = li.find("a", href=True)
        if not a_tag:
            continue

        title = a_tag.get_text(strip=True)
        href = a_tag["href"].strip()

        # 发布时间从 <span> 中提取
        span = li.find("span")
        pub_time = ""
        if span:
            span_text = span.get_text(strip=True)
            # 格式如 "2026.04.16 15:43:04 | 采购人：..."
            parts = span_text.split("|")
            if parts:
                pub_time = parts[0].strip()

        items.append((title, pub_time, href))
    return items


def extract_from_table(soup):
    """尝试从概要表格中提取字段"""
    data = {
        "项目名称": "",
        "预算金额": "",
        "项目联系人": "",
        "项目联系电话": "",
        "开标时间": "",
        "采购人名称": "",
        "采购人地址": "",
    }

    table = soup.find("div", class_="table")
    if not table:
        return data

    # 提取所有 tr
    for tr in table.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) < 2:
            continue

        label = tds[0].get_text(strip=True)
        value = tds[1].get_text(strip=True)

        if label == "采购项目名称" and not data["项目名称"]:
            data["项目名称"] = value
        elif label == "采购单位" and not data["项目名称"]:
            # 兜底：没有项目名称时，用采购单位
            data["项目名称"] = value
        elif label == "采购单位" and not data["采购人名称"]:
            data["采购人名称"] = value
        elif label == "采购单位地址" and not data["采购人地址"]:
            data["采购人地址"] = value
        elif label == "预算金额" and not data["预算金额"]:
            data["预算金额"] = value
        elif label == "项目联系人" and not data["项目联系人"]:
            data["项目联系人"] = value
        elif label == "项目联系电话" and not data["项目联系电话"]:
            data["项目联系电话"] = value
        elif label == "开标时间" and not data["开标时间"]:
            data["开标时间"] = value

    return data


def extract_from_content(soup):
    """从正文内容中用正则提取字段"""
    content_div = soup.find("div", class_="vF_detail_content")
    if not content_div:
        return {}

    text = content_div.get_text(separator="\n", strip=True)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    full_text = "\n".join(lines)

    result = {}

    # 项目编号：优先用较严格的模式匹配常见编号格式，失败再用宽松模式
    # 常见格式：字母、数字、横杠、下划线、点、括号等
    m = re.search(r"项目编号[：:]\s*([A-Za-z0-9\-_\.\(\)／/]+)", full_text)
    if m:
        val = m.group(1).strip()
        # 去掉末尾常见标点/污染字符
        val = re.sub(r"[\)、。，；\.]+$", "", val)
        result["项目编号"] = val
    else:
        m = re.search(r"项目编号[：:]\s*([^\n]{2,40})", full_text)
        if m:
            val = m.group(1).strip()
            # 去掉常见的后缀污染
            val = re.split(r"[）\)公开招标|竞争性谈判|询价|单一来源|成交|中标|公告]", val)[0]
            result["项目编号"] = val.strip()

    # 项目名称
    m = re.search(r"项目名称[：:]\s*([^\n]+)", full_text)
    if m:
        result["项目名称"] = m.group(1).strip()

    # 预算金额
    m = re.search(r"预算金额[：:]\s*([^\n]+)", full_text)
    if m:
        result["预算金额"] = m.group(1).strip()

    # 采购需求：尝试多种策略
    # 策略1：匹配 "采购需求：" 后到换行或特定截止词之间的内容
    m = re.search(r"采购需求[：:]\s*(.+?)(?=\n\s*(?:合同履行期限|最高限价|本项目不接受|二、申请人的资格要求))", full_text, re.DOTALL)
    if m:
        val = m.group(1).strip()
        # 如果包含大量表格文字或太短，尝试策略2
        if len(val.replace(" ", "").replace("\n", "")) < 10:
            m = None
    if not m:
        # 策略2：宽松匹配第一行（至少5个字符）
        m = re.search(r"采购需求[：:]\s*([^\n]{5,300})", full_text)
    if m:
        val = m.group(1).strip()
        # 清理多余换行和空格
        val = re.sub(r"\s+", " ", val)
        result["采购需求"] = val

    # 提交投标文件截止时间 / 开标时间
    # 先尝试精确匹配
    m = re.search(r"提交投标文件截止时间.*?[：:]\s*([^\n]+)", full_text)
    if m:
        result["提交投标文件截止时间"] = m.group(1).strip()
    else:
        # 尝试匹配 "并于 XXXX 前递交投标文件"
        m = re.search(r"并于\s*([^前\n]+?)\s*前递交投标文件", full_text)
        if m:
            result["提交投标文件截止时间"] = m.group(1).strip()
        else:
            # 尝试匹配 "开标时间"
            m = re.search(r"开标时间[：:]\s*([^\n]+)", full_text)
            if m:
                result["提交投标文件截止时间"] = m.group(1).strip()

    # 项目联系人
    m = re.search(r"项目联系人[：:]\s*([^\n]+)", full_text)
    if m:
        result["项目联系人"] = m.group(1).strip()

    # 电话（优先匹配 "电 话" 或 "联系电话"）
    m = re.search(r"电\s*话[：:]\s*([^\n]+)", full_text)
    if m:
        result["电话"] = m.group(1).strip()
    else:
        m = re.search(r"联系电话[：:]\s*([^\n]+)", full_text)
        if m:
            result["电话"] = m.group(1).strip()

    # 采购人名称（优先在"采购人信息"附近匹配"名 称"）
    m = re.search(r"(?:采购人信息|采购人)[^\n]*(?:\n.*?){0,3}名\s*称[：:]\s*([^\n]+)", full_text, re.DOTALL)
    if m:
        result["采购人名称"] = m.group(1).strip()
    else:
        # 兜底：匹配"采购单位"
        m = re.search(r"采购单位[：:]\s*([^\n]+)", full_text)
        if m:
            result["采购人名称"] = m.group(1).strip()

    # 采购人地址（优先在"采购人信息"附近匹配"地 址"）
    m = re.search(r"(?:采购人信息|采购人)[^\n]*(?:\n.*?){0,3}地\s*址[：:]\s*([^\n]+)", full_text, re.DOTALL)
    if m:
        result["采购人地址"] = m.group(1).strip()
    else:
        # 兜底：匹配通用"地址"
        m = re.search(r"地\s*址[：:]\s*([^\n]+)", full_text)
        if m:
            result["采购人地址"] = m.group(1).strip()

    return result


def parse_detail_page(html):
    """解析详情页，返回字典"""
    soup = BeautifulSoup(html, "html.parser")

    # 1. 尝试从表格提取
    table_data = extract_from_table(soup)

    # 2. 从正文提取
    content_data = extract_from_content(soup)

    # 3. 合并结果（正文优先级更高，表格作为兜底）
    result = {
        "项目编号": content_data.get("项目编号", ""),
        "项目名称": content_data.get("项目名称", table_data.get("项目名称", "")),
        "预算金额": content_data.get("预算金额", table_data.get("预算金额", "")),
        "采购需求": content_data.get("采购需求", ""),
        "提交投标文件截止时间": content_data.get("提交投标文件截止时间", table_data.get("开标时间", "")),
        "项目联系人": content_data.get("项目联系人", table_data.get("项目联系人", "")),
        "电话": content_data.get("电话", table_data.get("项目联系电话", "")),
        "采购人名称": content_data.get("采购人名称", table_data.get("采购人名称", "")),
        "采购人地址": content_data.get("采购人地址", table_data.get("采购人地址", "")),
    }

    # 清理 HTML 残留标签（极少情况）
    for k, v in result.items():
        if v:
            v = re.sub(r"<br\s*/?>", " ", v)
            v = re.sub(r"<[^>]+>", "", v)
            result[k] = v.strip()

    return result


def main():
    csv_filename = "ccgp_招标公告.csv"
    fieldnames = [
        "标题",
        "发布时间",
        "详情页链接",
        "项目编号",
        "项目名称",
        "预算金额",
        "采购需求",
        "提交投标文件截止时间",
        "项目联系人",
        "电话",
        "采购人名称",
        "采购人地址",
    ]

    with open(csv_filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        # 默认爬取前 5 页
        for page in range(1, 6):
            list_url = LIST_URL_TEMPLATE.format(
                page=page, kw=urllib.parse.quote(KEYWORD)
            )
            print(f"[→] 正在获取列表页第 {page} 页...")
            html = fetch_html(list_url)
            if not html:
                print(f"[!] 列表页第 {page} 页获取失败，跳过")
                continue

            items = parse_list_page(html)
            print(f"[*] 第 {page} 页共找到 {len(items)} 条公告")

            if not items:
                break

            for idx, (title, pub_time, detail_url) in enumerate(items, 1):
                print(f"  [{idx}/{len(items)}] {title[:40]}... -> {detail_url}")

                detail_html = fetch_html(detail_url)
                if not detail_html:
                    print(f"      [!] 详情页获取失败")
                    continue

                detail = parse_detail_page(detail_html)

                row = {
                    "标题": title,
                    "发布时间": pub_time,
                    "详情页链接": detail_url,
                    "项目编号": detail.get("项目编号", ""),
                    "项目名称": detail.get("项目名称", ""),
                    "预算金额": detail.get("预算金额", ""),
                    "采购需求": detail.get("采购需求", ""),
                    "提交投标文件截止时间": detail.get("提交投标文件截止时间", ""),
                    "项目联系人": detail.get("项目联系人", ""),
                    "电话": detail.get("电话", ""),
                    "采购人名称": detail.get("采购人名称", ""),
                    "采购人地址": detail.get("采购人地址", ""),
                }
                writer.writerow(row)

                # 每访问一个详情页等待 1 秒，避免请求过快
                time.sleep(1)

            # 列表页之间也稍作等待
            time.sleep(2)

    print(f"\n[√] 数据已保存到: {csv_filename}")


if __name__ == "__main__":
    main()
