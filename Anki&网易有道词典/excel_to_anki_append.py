#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
把 youdao_defs.xlsx（两列：word, definition）追加导入到现有 Anki 牌组
新增：导入前总览（N/A/R）+ 交互确认
依赖：pip install pandas openpyxl requests
确保：Anki 已运行，安装 AnkiConnect（http://localhost:8765）
"""

import requests
import pandas as pd
from typing import List, Dict, Any

ANKI_URL = "http://localhost:8765"
MODEL_NAME = "Youdao Basic (Auto)"     # 若不存在会自动创建
TAGS = ["youdao", "auto"]
BATCH_SIZE = 100                        # 分批提交大小（长名单更稳）


# ---------- 基础 RPC ----------
def _post(payload: Dict[str, Any]) -> Dict[str, Any]:
    r = requests.post(ANKI_URL, json=payload)
    r.raise_for_status()
    return r.json()

def invoke(action: str, **params):
    payload = {"action": action, "version": 6, "params": params}
    resp = _post(payload)
    # 某些版本会把重复等信息塞进 error；此处宽容返回
    if resp.get("error"):
        return {"error": resp["error"], "result": resp.get("result")}
    return resp.get("result")


# ---------- 准备环境 ----------
def ensure_deck(deck_name: str):
    decks = invoke("deckNames")
    if isinstance(decks, dict):  # 兼容 error 透传
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
    res = invoke(
        "createModel",
        modelName=model_name,
        inOrderFields=[f["name"] for f in fields],
        css=css,
        isCloze=False,
        cardTemplates=[{"Name": t["Name"], "Front": t["Front"], "Back": t["Back"]} for t in templates],
    )
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

    ensure_deck(deck_name)
    ensure_model(MODEL_NAME)

    # 构造 notes（此处不直接 addNotes，先 canAddNotes 预检）
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
                "duplicateScope": "deck"
            }
        }
        words.append(word)
        notes.append(note)

    if not notes:
        print("没有可导入的记录。")
        return

    # ===== 预检：canAddNotes =====
    print("\n正在预检可添加性（canAddNotes）…")
    can = invoke("canAddNotes", notes=notes)
    # 若 can 返回异常，退化为全部尝试添加；但仍给出警告
    if isinstance(can, dict) and can.get("error"):
        print(f"[警告] canAddNotes 返回异常，将在导入阶段再判断重复：{can['error']}")
        addable_mask = [True] * len(notes)
    else:
        addable_mask = list(can)

    total = len(notes)
    addable = sum(1 for x in addable_mask if x)
    not_addable = total - addable

    # 列出前 20 个预计重复/不可加的单词做预览
    preview_dups = [w for w, ok in zip(words, addable_mask) if not ok][:20]

    print("\n====== 导入前总览 ======")
    print(f"总记录：{total}")
    print(f"可新增：{addable}")
    print(f"预计重复/不可添加：{not_addable}")
    if preview_dups:
        print("预计重复（前 20 个预览）:")
        for w in preview_dups:
            print(f"  - {w}")
    print("========================")

    # 交互确认
    go = input(f"\n是否继续导入可新增的 {addable} 条？[y/N]: ").strip().lower()
    if go not in ("y", "yes"):
        print("已取消导入。")
        return

    # ===== 导入阶段 =====
    added_total = 0
    skipped_total = 0
    failed_total = 0

    if isinstance(can, dict) and can.get("error"):
        # 退化路径：不能用 canAddNotes，只能直接提交并根据返回判断
        print("\n[提示] 进入退化导入路径（无法使用 canAddNotes 结果）…")
        for batch in chunked(list(zip(words, notes)), BATCH_SIZE):
            batch_words = [w for w, _ in batch]
            batch_notes = [n for _, n in batch]
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
    else:
        # 正常路径：只导入 addable 的条目
        pairs_add = [(w, n) for (w, n), ok in zip(zip(words, notes), addable_mask) if ok]
        pairs_skip = [(w, n) for (w, n), ok in zip(zip(words, notes), addable_mask) if not ok]

        # 显式打印预计跳过的（重复/不可添加）
        for w, _ in pairs_skip:
            print(f"[重复跳过] {w}")

        for batch in chunked(pairs_add, BATCH_SIZE):
            batch_words = [w for w, _ in batch]
            batch_notes = [n for _, n in batch]
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

    print(f"\n完成：新增 {added_total} 条，跳过/重复 {skipped_total} 条，失败 {failed_total} 条，总计 {total} 条。")


if __name__ == "__main__":
    main()
