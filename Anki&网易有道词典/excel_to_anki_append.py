#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
把 youdao_defs.xlsx（两列：word, definition）追加导入到现有 Anki 牌组
改进：
  - 导入前：先对 Excel 内部去重（大小写无关）
  - 过滤：排除牌组里已存在的 Front（从 Anki 拉现有 Front）
  - 预检总览 + 交互确认
  - 分批导入 + 逐条日志
依赖：pip install pandas openpyxl requests
确保：Anki 运行且已安装 AnkiConnect（http://localhost:8765）
"""

import requests
import pandas as pd
from typing import List, Dict, Any, Set

ANKI_URL = "http://localhost:8765"
MODEL_NAME = "Youdao Basic (Auto)"
TAGS = ["youdao", "auto"]
BATCH_SIZE = 50  # 稳一点


# ---------- 基础 RPC ----------
def _post(payload: Dict[str, Any]) -> Dict[str, Any]:
    r = requests.post(ANKI_URL, json=payload)
    r.raise_for_status()
    return r.json()

def invoke(action: str, **params):
    payload = {"action": action, "version": 6, "params": params}
    resp = _post(payload)
    if resp.get("error"):
        return {"error": resp["error"], "result": resp.get("result")}
    return resp.get("result")


# ---------- 环境 ----------
def ensure_deck(deck_name: str):
    decks = invoke("deckNames")
    if isinstance(decks, dict):
        raise RuntimeError(f"AnkiConnect error (deckNames): {decks}")
    if deck_name not in decks:
        res = invoke("createDeck", deck=deck_name)
        if isinstance(res, dict) and res.get("error"):
            raise RuntimeError(f"AnkiConnect error (createDeck): {res}")

def ensure_model(model_name: str):
    models = invoke("modelNames")
    if isinstance(models, dict):
        raise RuntimeError(f"AnkiConnect error (modelNames): {models}")
    if model_name in models:
        return
    templates = [{
        "Name": "Card 1",
        "Front": "{{Front}}",
        "Back": "{{Front}}<hr id=answer>{{Back}}"
    }]
    css = """
.card { font-family: -apple-system, Segoe UI, Roboto, Noto Sans SC, Arial; font-size: 20px; line-height: 1.5; }
hr { margin: 12px 0; }
"""
    fields = [{"name": "Front"}, {"name": "Back"}]
    res = invoke("createModel",
                 modelName=model_name,
                 inOrderFields=[f["name"] for f in fields],
                 css=css,
                 isCloze=False,
                 cardTemplates=[{"Name": t["Name"], "Front": t["Front"], "Back": t["Back"]} for t in templates])
    if isinstance(res, dict) and res.get("error"):
        raise RuntimeError(f"AnkiConnect error (createModel): {res}")


# ---------- 工具 ----------
def newline_to_html(s: str) -> str:
    if s is None:
        return ""
    return str(s).replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")

def chunked(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i + size]

def get_existing_fronts(deck_name: str) -> Set[str]:
    """获取目标牌组里现有 Front 字段（小写）"""
    # 注意 deck 名含空格/中文时建议加引号
    query = f'deck:"{deck_name}"'
    note_ids = invoke("findNotes", query=query)
    if isinstance(note_ids, dict) and note_ids.get("error"):
        raise RuntimeError(f"AnkiConnect error (findNotes): {note_ids}")
    if not note_ids:
        return set()
    info = invoke("notesInfo", notes=note_ids)
    if isinstance(info, dict) and info.get("error"):
        raise RuntimeError(f"AnkiConnect error (notesInfo): {info}")
    fronts = set()
    for it in info:
        fields = it.get("fields", {})
        front = ""
        if "Front" in fields:
            front = fields["Front"]["value"]
        elif "front" in fields:
            front = fields["front"]["value"]
        fronts.add(front.strip().lower())
    return fronts


# ---------- 主流程 ----------
def main():
    xlsx_path = input("请输入 youdao_defs.xlsx 文件路径: ").strip().strip('"').strip("'")
    deck_name = input("请输入要追加导入的 Anki 牌组名称: ").strip()

    # 读 Excel（兼容大小写/不规范列名）
    df = pd.read_excel(xlsx_path, dtype=str).fillna("")
    lower_cols = {c.lower(): c for c in df.columns}
    if "word" not in lower_cols or "definition" not in lower_cols:
        raise RuntimeError("Excel 需要包含列名：word, definition")
    df = df.rename(columns={lower_cols["word"]: "word", lower_cols["definition"]: "definition"})

    # 1) 先去除 Excel 内部重复（大小写无关，保留首次出现）
    before = len(df)
    df["_key"] = df["word"].astype(str).str.strip().str.lower()
    df = df[~df["_key"].duplicated(keep="first")].copy()
    removed_in_file = before - len(df)
    df.drop(columns=["_key"], inplace=True)

    ensure_deck(deck_name)
    ensure_model(MODEL_NAME)

    # 2) 过滤掉牌组里已存在的 Front
    existing = get_existing_fronts(deck_name)
    df["_key"] = df["word"].astype(str).str.strip().str.lower()
    df = df[~df["_key"].isin(existing)].copy()
    removed_in_deck = before - removed_in_file - len(df)
    df.drop(columns=["_key"], inplace=True)

    # 构造 notes
    words: List[str] = []
    notes: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        word = (row.get("word") or "").strip()
        definition = (row.get("definition") or "").strip()
        if not word:
            continue
        back_html = newline_to_html(definition)
        note = {
            "deckName": deck_name,
            "modelName": MODEL_NAME,
            "fields": {"Front": word, "Back": back_html},
            "tags": TAGS,
            "options": {
                "allowDuplicate": False,
                "duplicateScope": "deck",
                "duplicateScopeOptions": {
                    "deckName": deck_name,
                    "checkChildren": False,
                    "checkAllModels": False
                }
            }
        }
        words.append(word)
        notes.append(note)

    if not notes:
        print("没有可导入的记录。")
        print(f"(Excel 内部去重移除 {removed_in_file} 条；牌组中已存在移除 {removed_in_deck} 条)")
        return

    # 3) 预检总览
    total = len(notes)
    can = invoke("canAddNotes", notes=notes)
    if isinstance(can, dict) and can.get("error"):
        print(f"[警告] canAddNotes 异常，将直接导入并逐条判断：{can['error']}")
        addable_mask = [True] * total
        predicted_skips = 0
    else:
        addable_mask = list(can)
        predicted_skips = addable_mask.count(False)

    print("\n====== 导入前总览 ======")
    print(f"Excel 原始：{before}")
    print(f"Excel 内部去重移除：{removed_in_file}")
    print(f"牌组已存在移除：{removed_in_deck}")
    print(f"待提交总数：{total}")
    print(f"预测可新增：{addable_mask.count(True)}")
    print(f"预测重复/不可添加：{predicted_skips}")
    preview_dups = [w for w, ok in zip(words, addable_mask) if not ok][:20]
    if preview_dups:
        print("预计重复（前 20 个）：")
        for w in preview_dups:
            print(f"  - {w}")
    print("========================")

    go = input(f"\n是否继续导入可新增的 {addable_mask.count(True)} 条？[y/N]: ").strip().lower()
    if go not in ("y", "yes"):
        print("已取消导入。")
        return

    # 4) 导入（分批 + 逐条日志）
    def add_batch(batch_words, batch_notes):
        nonlocal added_total, skipped_total, failed_total
        add_res = invoke("addNotes", notes=batch_notes)
        if isinstance(add_res, dict) and add_res.get("error"):
            errs = add_res["error"]
            if isinstance(errs, list):
                for w, msg in zip(batch_words, errs):
                    if "duplicate" in str(msg).lower():
                        print(f"[重复跳过] {w}")
                        skipped_total += 1
                    else:
                        print(f"[失败] {w} -> {msg}")
                        failed_total += 1
            else:
                print(f"[失败] 整批失败 -> {errs}")
                failed_total += len(batch_words)
        else:
            for w, rid in zip(batch_words, add_res):
                if isinstance(rid, int):
                    print(f"[导入成功] {w} -> noteId={rid}")
                    added_total += 1
                elif rid is None:
                    print(f"[重复跳过] {w}")
                    skipped_total += 1
                else:
                    print(f"[失败] {w} -> {rid}")
                    failed_total += 1

    added_total = skipped_total = failed_total = 0

    pairs_add = [(w, n) for (w, n), ok in zip(zip(words, notes), addable_mask) if ok]
    pairs_skip = [(w, n) for (w, n), ok in zip(zip(words, notes), addable_mask) if not ok]

    for w, _ in pairs_skip:
        print(f"[重复跳过] {w}")

    for batch in chunked(pairs_add, BATCH_SIZE):
        batch_words = [w for w, _ in batch]
        batch_notes = [n for _, n in batch]
        add_batch(batch_words, batch_notes)

    print(f"\n完成：新增 {added_total} 条，跳过/重复 {skipped_total} 条，失败 {failed_total} 条，总计 {total} 条。")


if __name__ == "__main__":
    main()
