# Anki & 网易有道词典

这一子模块包含两个 Python 脚本，配合使用可以实现：  
1. 批量抓取 **有道词典** 的单词释义  
2. 将结果追加导入到 **Anki** 现有牌组  

---

## 功能介绍

- **youdao_to_excel.py**  
  - 输入：`words.txt`，每行一个英文单词  
  - 动作：访问有道词典网页，抓取释义  
  - 输出：`youdao_defs.xlsx`（两列：`word` 和 `definition`，释义多行显示）

- **excel_to_anki_append.py**  
  - 输入：`youdao_defs.xlsx`  
  - 动作：通过 [AnkiConnect](https://ankiweb.net/shared/info/2055492159) 把单词与释义追加到指定牌组  
  - 输出：单词卡（Front=word，Back=definition，支持换行显示）  

---

## 环境依赖

```bash
pip install requests beautifulsoup4 lxml pandas openpyxl
并且需要：

已安装 Anki，保持运行

已安装插件 AnkiConnect（在 Anki → 工具 → 插件 → 获取插件，输入 2055492159，重启 Anki）

验证成功的方法：在浏览器访问 http://localhost:8765，看到：

json
Copy code
{"apiVersion": "AnkiConnect v.6"}
说明接口已启用。

使用方法
1. 抓取有道释义 → Excel
创建一个 words.txt，例如：

nginx
Copy code
ephemeral
resilient
meticulous
运行：

bash
Copy code
python youdao_to_excel.py
脚本会要求输入 words.txt 的完整路径，例如：

mathematica
Copy code
请输入 words.txt 文件路径: C:\Users\Administrator\Desktop\testing1\words.txt
在相同目录下生成 youdao_defs.xlsx，每个单词一行，释义多行显示。

2. Excel → 追加导入 Anki
确保 Anki 打开，AnkiConnect 插件已启用。

运行：

bash
Copy code
python excel_to_anki_append.py
输入：

mathematica
Copy code
请输入 youdao_defs.xlsx 文件路径: C:\Users\Administrator\Desktop\testing1\youdao_defs.xlsx
请输入要追加导入的 Anki 牌组名称: My Vocabulary
日志会逐条显示：

python-repl
Copy code
[导入成功] ephemeral -> noteId=1623456789012
[重复跳过] resilient
...
完成：新增 20 条，跳过/重复 5 条，总计 25 条。
特性
批次内查重：words.txt 中重复单词自动跳过

释义换行：Excel 中的多条释义用 \n 分隔，Anki 导入时勾选 Keep line breaks 即可显示为多行

避免重复导入：在同一 Deck 中，已有的 Front（单词）不会重复添加

清晰日志：实时显示导入状态（成功 / 跳过）
