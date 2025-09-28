#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
批量抓取有道词典释义 -> 导出 Excel（两列：word, definition）
改进点：
  - 运行时输入 words.txt 路径（自动清理不可见字符）
  - 批次内查重（大小写无关）并跳过重复单词
  - 输出 youdao_defs.xlsx 到 words.txt 的同一路径
  - 多条释义用换行 \n 连接（Excel 单元格多行；Anki 导入勾选“Keep line breaks”即可换行显示）
依赖：pip install requests beautifulsoup4 lxml pandas openpyxl
"""

import time
import random
import pathlib
import logging
from typing import List, Tuple

import requests
from bs4 import BeautifulSoup
import pandas as pd

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
BASE_URL = "https://dict.youdao.com/w/eng/{word}/"

SLEEP_RANGE = (0.8, 1.6)  # 请求间隔
TIMEOUT = 12
RETRY = 2

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")


def clean_invisibles(s: str) -> str:
    # 清理常见不可见字符：U+202A/U+202B（方向嵌入）、BOM 等
    return s.replace("\u202a", "").replace("\u202b", "").replace("\ufeff", "")


def load_words(path: pathlib.Path) -> List[str]:
    if not path.exists():
        raise FileNotFoundError(f"未找到 {path}，请检查路径是否正确")
    words: List[str] = []
    seen_ci = set()  # case-insensitive 去重
    for line in path.read_text(encoding="utf-8").splitlines():
        w = clean_invisibles(line).strip()
        if not w or w.startswith("#"):
            continue
        key = w.lower()
        if key in seen_ci:
            continue  # 批次内重复 => 跳过
        seen_ci.add(key)
        words.append(w)
    return words


def fetch_html(url: str) -> str:
    headers = {"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"}
    for attempt in range(RETRY + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=TIMEOUT)
            if resp.status_code == 200:
                return resp.text
            logging.warning(f"HTTP {resp.status_code} for {url}")
        except requests.RequestException as e:
            logging.warning(f"请求失败 {e} (尝试 {attempt+1}/{RETRY+1})")
        time.sleep(0.7 + 0.5 * attempt)
    return ""


def parse_definitions(html: str) -> List[str]:
    """
    优先抓“基本释义”块；若缺失，兜底抓其他区域的简短文本。
    返回多条释义的列表（不含分隔符）。
    """
    if not html:
        return []
    soup = BeautifulSoup(html, "lxml")

    defs = []
    phrs = soup.select_one("#phrsListTab")
    if phrs:
        for li in phrs.select("div.trans-container ul li"):
            text = " ".join(li.get_text(" ", strip=True).split())
            if text:
                defs.append(text)

    if not defs:
        for blk in soup.select("div.trans-container ul li"):
            text = " ".join(blk.get_text(" ", strip=True).split())
            if text:
                defs.append(text)

    if not defs:
        collins_items = soup.select("div#collinsResult div.collinsMajorTrans p")
        for p in collins_items[:3]:
            text = " ".join(p.get_text(" ", strip=True).split())
            if text:
                defs.append(text)

    # 去重 & 截断
    seen, uniq = set(), []
    for d in defs:
        if d not in seen:
            seen.add(d)
            uniq.append(d)
    return uniq[:5]  # 最多取前 5 条


def get_youdao_definition(word: str) -> str:
    url = BASE_URL.format(word=word)
    html = fetch_html(url)
    defs = parse_definitions(html)
    if not defs:
        return ""
    # 用换行连接，而不是用 |
    return "\n".join(defs)


def main():
    # 读取并清理输入路径
    raw = input("请输入 words.txt 文件路径: ").strip().strip('"').strip("'")
    raw = clean_invisibles(raw)
    words_path = pathlib.Path(raw)
    print(f"实际使用路径: {words_path}")

    words = load_words(words_path)

    # 输出文件放在 words.txt 同目录
    output_xlsx = words_path.parent / "youdao_defs.xlsx"

    rows: List[Tuple[str, str]] = []
    for i, w in enumerate(words, 1):
        time.sleep(random.uniform(*SLEEP_RANGE))
        definition = get_youdao_definition(w)
        # 如果本词抓到的释义里自己含有竖线，顺手替换为换行，保持干净
        definition = definition.replace(" | ", "\n").replace("|", "\n")
        rows.append((w, definition))
        preview = definition.replace("\n", " \\n ")
        logging.info(f"[{i}/{len(words)}] {w} -> {preview[:100]}{'...' if len(preview)>100 else ''}")

    df = pd.DataFrame(rows, columns=["word", "definition"])
    # 保留换行：Excel里同一个单元格会显示为多行（Alt+Enter）
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    logging.info(f"已生成 {output_xlsx}（两列：word, definition，释义为多行）")
    print("\n导入 Anki 提示：\n- 使用“文件→导入”，选择该 Excel 转成的 TSV/CSV 或直接用 CSV 导出\n- 导入向导里勾选 ‘保留换行(Keep line breaks)’，即可在卡片上按行显示\n- 字段映射：Front=word, Back=definition\n")


if __name__ == "__main__":
    main()
