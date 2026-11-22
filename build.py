#!/usr/bin/env python3
"""把 src/gas/ 裡的 Apps Script 程式碼嵌進 src/site.html，輸出 index.html。

教學網頁上讓學員複製的程式碼，和 test.js 測試的是同一份檔案。
用法：python3 build.py
"""
import json
from pathlib import Path

ROOT = Path(__file__).parent
FILES = {"Capture": "Capture.gs", "Web": "Web.gs", "Index": "Index.html",
         "Ask": "Ask.gs", "Weekly": "Weekly.gs", "Compose": "Compose.gs", "Secretary": "Secretary.gs"}

gas = {k: (ROOT / "src" / "gas" / f).read_text(encoding="utf-8") for k, f in FILES.items()}
# 放在 <script type="application/json"> 裡：把 < 換成 <，避免程式碼裡的 </script> 把標籤提早關掉
blob = json.dumps(gas, ensure_ascii=False).replace("<", "\\u003c")
site = (ROOT / "src" / "site.html").read_text(encoding="utf-8")
assert site.count("__GAS_JSON__") == 1
(ROOT / "index.html").write_text(site.replace("__GAS_JSON__", blob), encoding="utf-8")
print("index.html 已產生，%d KB" % (len((ROOT / "index.html").read_bytes()) // 1024))
