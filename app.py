"""Operation: Streamlit UI with Supabase Auth and authenticated shared storage."""
import json
import os
import time
import uuid
import hashlib
import urllib.error
import urllib.request
from pathlib import Path

import streamlit as st
# Read cached worksheet values without running workbook macros or formulas.
import io
import re
import zipfile
import posixpath
from datetime import datetime, timedelta, timezone
from xml.etree import ElementTree as ET


def parse_order_workbook(blob, filename, source_url=""):
    if len(blob) > 15 * 1024 * 1024:
        raise ValueError("15MB 이하의 엑셀 파일을 선택해주세요.")
    if not zipfile.is_zipfile(io.BytesIO(blob)):
        raise ValueError("올바른 xlsx 또는 xlsm 파일이 아니에요.")
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        if sum(x.file_size for x in z.infolist()) > 100 * 1024 * 1024:
            raise ValueError("엑셀 내부 데이터가 너무 큽니다.")
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheet = next((s for s in wb.findall("s:sheets/s:sheet", ns)
                      if s.attrib.get("name", "").strip().lower() == "open so status"), None)
        if sheet is None:
            raise ValueError("Open SO Status 시트를 찾지 못했어요.")
        rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = next(r.attrib["Target"] for r in rels if r.attrib["Id"] == rid)
        path = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            strings = ["".join(si.itertext()) for si in ET.fromstring(z.read("xl/sharedStrings.xml"))]
        rows = []
        for row in ET.fromstring(z.read(path)).findall("s:sheetData/s:row", ns):
            cells = {}
            for c in row.findall("s:c", ns):
                col = re.sub(r"\d", "", c.attrib["r"])
                value = c.findtext("s:v", default="", namespaces=ns)
                kind = c.attrib.get("t")
                if kind == "s":
                    value = strings[int(value)] if value else ""
                elif kind == "inlineStr":
                    value = "".join(c.find("s:is", ns).itertext())
                cells[col] = value.strip()
            rows.append((int(row.attrib["r"]), cells))
        header = next(((n, r) for n, r in rows if r.get("C", "").upper() == "STATUS"
                       and r.get("H", "").upper() == "CUSTOMER" and r.get("J", "").upper() == "ORDER VALUE"), None)
        if not header:
            raise ValueError("Status / Customer / Order Value 열 구성이 달라요. 기존 요약은 유지됩니다.")
        props = wb.find("s:workbookPr", ns)
        epoch = datetime(1904, 1, 1) if props is not None and props.attrib.get("date1904") in ("1", "true") else datetime(1899, 12, 30)
        def date(value):
            try:
                return (epoch + timedelta(days=float(value))).date().isoformat() if value else ""
            except (ValueError, OverflowError):
                return value
        orders, excluded, ids = [], {}, set()
        for n, r in rows:
            if n <= header[0] or not (r.get("B") or r.get("H") or r.get("I")):
                continue
            status = re.sub(r"\s+", " ", r.get("C", "").upper())
            if status.startswith("#"):
                raise ValueError(f"{n}행의 상태에 엑셀 오류가 있어요.")
            if "SENT" in status or status in ("CANCELLED", "CANCELED"):
                excluded[status] = excluded.get(status, 0) + 1
                continue
            if not status or not r.get("H"):
                raise ValueError(f"{n}행에 상태 또는 회사명이 없어 집계하지 못했어요.")
            key = (r.get("H"), r.get("B"), r.get("I"))
            if key in ids:
                raise ValueError(f"{n}행의 PO/SO가 중복돼요. 중복 여부를 확인해주세요.")
            ids.add(key)
            try:
                amount = round(float(r["J"]), 2) if r.get("J") else None
                if amount is not None and not (-1e12 < amount < 1e12):
                    raise ValueError()
            except ValueError:
                raise ValueError(f"{n}행의 금액을 읽을 수 없어요.")
            orders.append({"row": n, "status": status, "customer": r["H"],
                           "po": r.get("B", ""), "so": r.get("I", ""), "amount": amount,
                           "poDate": date(r.get("D", "")), "readyDate": date(r.get("E", "")),
                           "shipFrom": r.get("G", ""), "note": r.get("O", "")})
        return {"version": 1, "orders": orders, "excluded": excluded,
                "sourceFile": filename, "sourceSheet": "Open SO Status", "sourceUrl": source_url,
                "checkedAt": datetime.now(timezone.utc).isoformat(), "sourceRows": len(rows) - header[0]}


HERE = Path(__file__).parent
st.set_page_config(page_title="Operation", page_icon="📋", layout="wide")


def setting(name, default=""):
    if os.environ.get(name):
        return os.environ[name]
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


URL = setting("SUPABASE_URL").rstrip("/")
KEY = setting("SUPABASE_PUBLISHABLE_KEY")


def api(path, method="GET", payload=None, token=None):
    headers = {"apikey": KEY, "Content-Type": "application/json"}
    if token:
        headers["Authorization"] = "Bearer " + token
    request = urllib.request.Request(
        URL + path, headers=headers, method=method,
        data=None if payload is None else json.dumps(payload).encode(),
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            return json.load(response)
    except urllib.error.HTTPError as error:
        try:
            detail = json.loads(error.read())
            message = detail.get("msg") or detail.get("message") or ""
        except Exception:
            message = ""
        raise RuntimeError(message or f"서버 요청 실패 ({error.code})") from None


def session_token():
    session = st.session_state.get("auth")
    if not session:
        return None
    if session.get("expires_at", 0) < time.time() + 60:
        session = api("/auth/v1/token?grant_type=refresh_token", "POST", {
            "refresh_token": session["refresh_token"],
        })
        session.setdefault("expires_at", time.time() + session.get("expires_in", 3600))
        st.session_state.auth = session
    return session["access_token"]


BOARD_HTML = '<div id="operation-v2">\n<style>\n#operation-v2{color:var(--foreground);font-family:inherit}#operation-v2 section{margin-top:22px}#operation-v2 .fields{display:grid;grid-template-columns:minmax(140px,2fr) minmax(100px,1fr) minmax(125px,1fr);gap:12px}#operation-v2 .line{padding:12px 0;border-bottom:1px solid var(--border);display:flex;gap:12px;align-items:flex-start}#operation-v2 .grow{flex:1;min-width:0;overflow-wrap:anywhere}#operation-v2 .secondary{color:var(--muted-foreground)}#operation-v2 .actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}#operation-v2 .note{white-space:pre-wrap;overflow-wrap:anywhere}#operation-v2 .details{padding:14px 0}#operation-v2 .result{overflow-wrap:anywhere}#operation-v2 input[type=checkbox]{accent-color:var(--primary);width:18px;height:18px}#operation-v2 summary{padding:8px 0}#operation-v2 .head{display:flex;gap:12px;flex-wrap:wrap;align-items:center;padding-right:90px}@media(max-width:550px){#operation-v2 .fields{grid-template-columns:1fr}#operation-v2 .head{padding-right:0}}\n</style>\n<style>#operation-v2 .section-move{display:inline-block;width:115px}#operation-v2 .section-move select{max-width:100%}#operation-v2 [data-section-history]{max-width:100%}#operation-v2 [data-section-history][open]{flex-basis:100%}</style>\n<style>#operation-v2 .daily-grid{grid-template-columns:minmax(100px,1fr) 55px 55px 120px}#operation-v2 .daily-header>span:nth-child(2),#operation-v2 .daily-header>span:nth-child(3){text-align:center}#operation-v2 .progress-check{display:flex;justify-content:center;align-items:center}@media(max-width:450px){#operation-v2 .daily-grid{grid-template-columns:minmax(85px,1fr) 42px 42px 110px}}</style>\n<style>#operation-v2 .daily-grid{display:grid;grid-template-columns:minmax(100px,1fr) 55px 55px 120px;gap:10px;align-items:center;padding:10px 0;border-bottom:1px solid var(--border)}#operation-v2 .daily-header{border-bottom:0}#operation-v2 .progress-check{text-align:center}#operation-v2 .daily-grid .inline-slot{grid-column:1/-1}@media(max-width:450px){#operation-v2 .daily-grid{grid-template-columns:minmax(85px,1fr) 42px 42px 110px;gap:4px}}</style>\n<style>#operation-v2 .manage-row{display:flex;align-items:center;gap:8px;flex-wrap:nowrap}#operation-v2 .manage-row .actions{display:contents}#operation-v2 .manage-row>*{flex-shrink:0}#operation-v2 .manage-row button{white-space:nowrap}#operation-v2 .date-cell{position:relative;display:inline-flex;align-items:center;min-width:52px;min-height:32px;white-space:nowrap}#operation-v2 .date-cell input{position:absolute;inset:0;width:100%;height:100%;opacity:0}#operation-v2 .date-cell:focus-within{outline:2px solid var(--ring)}#operation-v2 [data-item-row]>td{white-space:nowrap}#operation-v2 [data-inline-editor]>td{white-space:normal}#operation-v2 .title-cell{display:flex;align-items:center;gap:4px;min-width:360px}</style>\n<style>#operation-v2 .daily-row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;padding:8px 0;border-bottom:1px solid var(--border)}#operation-v2 .inline-slot:empty{display:none}#operation-v2 .inline-slot{flex-basis:100%}#operation-v2 [hidden]{display:none!important}</style>\n<style>#operation-v2 .task-name{display:block;width:220px;flex:0 0 220px}#operation-v2 .manage-row{gap:4px}</style><div class="head"><strong id="date-label"></strong><button id="next-day" class="btn" type="button">다음 날 보기</button><button id="real-day" class="btn" type="button" hidden>오늘로</button><button id="news-toggle" class="btn" type="button">🔔 새 소식</button></div>\n<div class="viz-row" style="align-items:center"><label for="my-emoji">내 이모지</label><select id="my-emoji" aria-label="내 이모지"></select><span id="my-profile"></span></div>\n<section id="news-panel" hidden><h3>새 소식</h3><div id="event-list"></div></section>\n<div id="notice" role="status" aria-live="polite"></div>\n<section id="review-panel" hidden><h3 id="review-title">새 노트와 결과물 확인 필요</h3><div id="review-list"></div></section>\n<section><details><summary class="cursor-interaction">＋ 새 업무 추가</summary><form id="add-form"><div class="fields"><label class="form-label">넣을 섹션<select id="task-section" class="form-select"></select></label><label id="new-section-label" class="form-label" hidden>새 섹션 이름<input id="task-new-section" class="form-control" maxlength="40" placeholder="예: 출시 준비"></label><label class="form-label">할 일<input id="task-title" required maxlength="100" class="form-control"></label><label class="form-label">담당<select id="task-owner" class="form-select"><option>미정</option><option>Hugo</option><option>제이나</option><option>함께</option></select></label><label class="form-label">마감일<input id="task-due" class="form-control" type="date"></label></div><div class="actions"><label class="form-check"><input id="task-daily" class="form-check-input" type="checkbox"><span class="form-check-label">매일 반복</span></label><button class="btn btn-primary" type="submit">추가하기</button><button id="add-close" class="btn" type="button" hidden>닫기</button></div></form></details></section>\n<section><div class="viz-row" style="align-items:center"><h3 id="new-title">오늘 새로 추가된 일</h3><button id="new-toggle" type="button" aria-expanded="true" aria-controls="new-content">▾ 접기</button></div><div id="new-content"><div id="new-list"></div><div id="added-dates" class="viz-row" aria-label="추가 날짜별 보기"></div></div></section>\n<hr>\n<section><div class="viz-row" style="align-items:center"><h3>Daily 할 일</h3><button id="daily-toggle" type="button" aria-expanded="true" aria-controls="daily-content">▾ 접기</button><button id="daily-add-toggle" type="button" aria-expanded="false" aria-controls="daily-add-form">＋ 추가</button></div><div id="daily-content"><form id="daily-add-form" hidden><div class="viz-row" style="align-items:center"><label class="form-label grow">매일 할 일<input id="daily-add-title" class="form-control" required maxlength="100" placeholder="매일 확인할 업무 이름"></label><button type="submit" class="btn btn-primary">Daily에 추가</button><button id="daily-add-close" type="button">닫기</button></div></form><div id="daily-list"></div></div></section>\n<hr>\n<section><h3>긴급 · 이번 주에 끝낼 일</h3><div class="secondary text-small">긴급도 높음·긴급, 이번 주 마감, 또는 ‘이번 주’ 체크한 업무가 함께 보여요.</div><div class="table-responsive"><table class="table table-sm"><thead><tr><th>업무</th><th>긴급도</th><th>중요도</th><th>담당</th><th>시작일</th><th>마감</th><th>관리</th></tr></thead><tbody id="focus-list"></tbody></table></div></section><hr><section><h3 id="active-title">전체 업무표</h3><div class="table-responsive"><table class="table table-sm"><tbody id="active-list"></tbody></table></div></section>\n<section id="deadline-section"><h3>마감일순</h3><div class="secondary text-small">마감일이 가까운 업무부터 표시돼요. 지난 마감일은 맨 위에 표시됩니다.</div><div class="table-responsive"><table class="table table-sm"><thead><tr><th>섹션</th><th>업무</th><th>긴급도</th><th>중요도</th><th>담당</th><th>시작일</th><th>마감</th><th>관리</th></tr></thead><tbody id="deadline-list"></tbody></table></div></section>\n<section id="orders-section"><hr><h3>진행 중 오더</h3><div id="orders-meta" class="secondary text-small"></div><div id="orders-summary" class="viz-row"></div><div id="orders-detail"></div></section>\n<section><details id="completed"><summary id="complete-title" class="cursor-interaction">완료함</summary><div id="complete-list"></div></details></section>\n<section><details><summary class="cursor-interaction">보관한 섹션</summary><div id="archived-sections"></div></details></section>\n<div id="deleted-list" hidden></div>\n<section id="editor" hidden><h3 id="edit-title">노트와 결과물</h3><form id="edit-form"><div class="fields"><label class="form-label">담당<select id="edit-owner" class="form-select"><option>미정</option><option>Hugo</option><option>제이나</option><option>함께</option></select></label><label class="form-label">시작일<input id="edit-start" type="date" class="form-control"></label><label class="form-label">마감일<input id="edit-due" type="date" class="form-control"></label></div><label class="form-label">진행 상황과 노트<textarea id="edit-note" rows="4" maxlength="600" class="form-control" placeholder="현재 상황, 부탁할 일, 결정한 내용을 적으세요"></textarea></label><label class="form-label">결과물 메모<textarea id="edit-result" rows="2" maxlength="300" class="form-control" placeholder="완료한 내용이나 결과물 설명"></textarea></label><label class="form-label">결과물 링크<input id="edit-link" type="url" maxlength="400" class="form-control" placeholder="https://..."></label><label class="form-label">이메일 제목<input id="edit-email" class="form-control" maxlength="200" placeholder="관련 이메일을 찾을 때 사용할 제목"></label><div class="actions"><button class="btn btn-primary" type="submit">저장</button><button class="btn" id="edit-close" type="button">닫기</button></div><div id="result-preview" class="result"></div></form></section>\n\n<section><details><summary class="cursor-interaction">Daily 체크 기록</summary><div id="daily-history"></div></details></section>\n</div>\n'
BOARD_CSS = '\n:host{--foreground:#f5f5f5;--muted-foreground:#999;--border:#383838;--primary:#4285f4;--ring:#4285f4;--blue:#4285f4;--purple:#b278ed;--green:#43c67b;--orange:#f3a547;--red:#fa646b;--yellow:#dccc5a;font:14px system-ui,sans-serif;color:var(--foreground)}\n*{box-sizing:border-box}h2{font-size:22px;margin:0}h3{font-size:18px;margin:14px 0}button,input,select,textarea{font:inherit;color:inherit}button,summary{cursor:pointer}button{border:1px solid #505050;border-radius:8px;background:#303030;padding:5px 10px;white-space:nowrap}button:hover{background:#414141}button:disabled,input:disabled{opacity:.45;cursor:default}input:not([type=checkbox]),textarea,select{background:#303030;border:1px solid #505050;border-radius:8px;padding:6px 8px;max-width:100%}input[type=checkbox]{accent-color:var(--primary)}.form-control{width:100%}.form-label{display:block;margin:8px 0}.form-label>input,.form-label>select,.form-label>textarea{display:block;margin-top:5px}.btn-primary{background:#2761c6}.btn-ghost{border-color:transparent;background:transparent}.table-responsive{overflow-x:auto}.table{border-collapse:collapse;width:100%;margin-top:10px}.table th,.table td{text-align:left;padding:8px 4px;border-bottom:1px solid var(--border)}.table th{font-size:13px}.table th:first-child{min-width:370px}.table th:not(:first-child){white-space:nowrap}.table input[type=text]{width:100%}.viz-row{display:flex;gap:6px;flex-wrap:wrap;margin-top:12px}.text-small{font-size:12px}hr{border:0;border-top:1px solid var(--border);margin:26px 0}#notice{color:#9ac1ff;margin-top:10px}#operation-v2 .daily-grid{grid-template-columns:minmax(200px,1fr) repeat(var(--daily-columns,2),55px) 120px}#operation-v2 .section-move{width:100px}#operation-v2 .fields{gap:8px}#operation-v2 .head{padding-right:0}a{color:#8cbbff}details{margin-top:10px}summary{color:#c6c6c6}\n\n#deadline-section .table th:first-child{min-width:130px}#deadline-section .deadline-area{display:block;white-space:nowrap}#deadline-section .table th:nth-child(2){min-width:370px}.priority-dot{font-weight:600;min-width:82px}\n\n.order-card{min-width:90px;display:flex;flex-direction:column;gap:6px;text-align:left;padding:12px}.order-card strong{font-size:22px}#orders-section .order-table th:first-child{min-width:150px}#orders-section .order-table td{vertical-align:top}#orders-section .order-table td:nth-child(2){white-space:nowrap}#orders-section .order-table td:last-child{min-width:200px;max-width:340px;white-space:normal}\n'
BOARD_JS = '\nexport default function(component){\nconst {parentElement,data,setTriggerValue}=component;\nif(parentElement._operation?.root===parentElement.querySelector(\'#operation-v2\')){parentElement._operation.receive(data);return;}\nlet revision=data.revision,pending=null;\nconst actor=data.actor;\nconst people=data.members.map(x=>x.display_name);\nconst bridge={widgetState:{modelContent:{board:data.board}},setWidgetState(payload){\n  return new Promise((resolve,reject)=>{\n    const id=crypto.randomUUID();pending={id,resolve,reject};\n    setTriggerValue(\'change\',{id,revision,board:payload.modelContent.board,message:payload.modelContent.board.events[0]?.text||\'업무판을 업데이트했어요\'});\n  });\n}};\n\nconst root=parentElement.querySelector(\'#operation-v2\'),q=id=>root.querySelector(\'#\'+id);\nconst editorNode=q(\'editor\');\nq(\'active-list\').closest(\'section\').before(q(\'deadline-section\'));\n\nconst noteDrafts=new Map();\nlet addedDay=null,addSection=null,editBase=null,orderStatus=null;\nconst addForm=q(\'add-form\'),addHome=addForm.parentElement;\nfunction mountAdd(){addHome.append(addForm);root.querySelectorAll(\'[data-inline-add]\').forEach(x=>x.remove());if(!addSection)return;const header=Array.from(root.querySelectorAll(\'[data-section]\')).find(x=>x.dataset.section===addSection);if(!header){addSection=null;return;}const row=document.createElement(\'tr\');row.dataset.inlineAdd=addSection;const cell=document.createElement(\'td\');cell.colSpan=7;cell.style.whiteSpace=\'normal\';row.append(cell);header.after(row);cell.append(addForm);q(\'add-close\').hidden=false;}\nfunction mountEditor(){if(!editId||editorNode.hidden)return;const button=root.querySelector(\'[data-item-row="\'+editId+\'"] [data-edit]\')||root.querySelector(\'[data-edit="\'+editId+\'"]\');if(!button)return;const row=button.closest(\'tr\');if(row){const next=document.createElement(\'tr\');next.dataset.inlineEditor=\'true\';const cell=document.createElement(\'td\');cell.colSpan=row.children.length;next.append(cell);row.after(next);cell.append(editorNode);}else{const daily=button.closest(\'.daily-grid\');if(daily)daily.querySelector(\'.inline-slot\').append(editorNode);else button.closest(\'.grow\').append(editorNode);}}\nconst day=()=>new Intl.DateTimeFormat(\'sv-SE\',{timeZone:\'America/New_York\',year:\'numeric\',month:\'2-digit\',day:\'2-digit\'}).format(new Date());\nconst plus=(d,n)=>{const x=new Date(d+\'T12:00:00Z\');x.setUTCDate(x.getUTCDate()+n);return x.toISOString().slice(0,10);};\nconst read=s=>s?.modelContent?.board;const valid=x=>x?.version===2&&Array.isArray(x.items)&&Array.isArray(x.events);\nlet state=structuredClone(data.board),viewDay=day(),editId=null,saving=false,dirtyEditor=false,seen=new Set(state.events.map(x=>x.id));\nlet drag=null;\nconst completions=new Set();\nlet completing=false;\nfunction completeNow(id){\n  if(data.readOnly||completions.has(id))return;\n  completions.add(id);render();notify(\'완료 처리 중…\');flushCompletions();\n}\nasync function flushCompletions(){\n  if(completing||!completions.size)return;\n  if(saving){setTimeout(flushCompletions,60);return;}\n  completing=true;\n  const ids=[...completions];\n  const ok=await mutate(()=>{for(const id of ids){const item=find(id);if(item){item.status=\'complete\';item.doneDate=day();}}},ids.length===1?(find(ids[0])?.title||\'업무\')+\' 완료했어요\':ids.length+\'개 업무를 완료했어요\');\n  ids.forEach(id=>completions.delete(id));completing=false;render();\n  if(!ok)notify(\'완료 저장에 실패해 업무를 다시 표시했어요. 다시 시도해주세요.\');\n  flushCompletions();\n}\nconst emojiChoices=[\'\',\'🙂\',\'😎\',\'😊\',\'🐻\',\'🐱\',\'🐶\',\'🦊\',\'🐼\',\'🐰\',\'🦁\',\'🐯\',\'🐸\',\'🐧\',\'🦉\',\'🌻\',\'🌷\',\'🍀\',\'⭐\',\'🌙\',\'🔥\',\'🚀\',\'💎\',\'🎯\'];\nfunction personLabel(name){const key=name===\'제이나\'?\'Jayna\':name;const emoji=state.memberEmoji?.[key];return (emojiChoices.includes(emoji)&&emoji?emoji+\' \':\'\')+name;}\nfunction renderOrders(){\n const summary=state.orderSummary;\n if(!summary){q(\'orders-meta\').textContent=\'원본 엑셀을 불러오면 상태별 진행 건수가 표시돼요.\';q(\'orders-summary\').innerHTML=\'\';q(\'orders-detail\').innerHTML=\'\';return;}\n const orders=summary.orders||[],statuses=[...new Set([\'PENDING\',\'PLANNED\',\'SUBMITTED\',\'CONFIRMED\',\'SHIPPED\',...orders.map(x=>x.status)])];\n const checked=new Date(summary.checkedAt),stale=Date.now()-checked.getTime()>36*60*60*1000;\n const stamp=Number.isNaN(checked.getTime())?\'확인 시각 없음\':checked.toLocaleString(\'ko-KR\',{timeZone:\'America/New_York\',month:\'2-digit\',day:\'2-digit\',hour:\'2-digit\',minute:\'2-digit\',hour12:false});\n const link=safeLink(summary.sourceUrl);\n q(\'orders-meta\').innerHTML=esc(\'원본 확인 \'+stamp+\' (뉴욕) · SENT / 취소 제외\')+(stale?\' <strong style="color:var(--orange)">최신 확인 필요</strong>\':\'\')+(link?\' · <a href="\'+esc(link)+\'" target="_blank" rel="noopener noreferrer">원본 엑셀 ↗</a>\':\'\');\n const cards=[[\'전체\',orders.length],...statuses.map(s=>[s,orders.filter(x=>x.status===s).length])];\n q(\'orders-summary\').innerHTML=cards.map(([s,n])=>\'<button type="button" class="order-card" data-order-status="\'+esc(s)+\'" aria-expanded="\'+(orderStatus===s)+\'" style="\'+(orderStatus===s?\'border-color:var(--blue)\':\'\')+\'"><span>\'+esc(s)+\'</span><strong>\'+n+\'건</strong></button>\').join(\'\');\n const list=orderStatus===\'전체\'?orders:orders.filter(x=>x.status===orderStatus);\n const money=n=>n===null||n===undefined?\'미정\':\'$\'+Number(n).toLocaleString(\'en-US\',{minimumFractionDigits:2,maximumFractionDigits:2});\n q(\'orders-detail\').innerHTML=orderStatus?\'<h4>\'+esc(orderStatus)+\' · \'+list.length+\'건</h4><div class="table-responsive"><table class="table order-table"><thead><tr><th>회사</th><th>금액</th><th>PO / SO</th><th>출고 예정</th><th>출고지</th><th>상태 / 메모</th></tr></thead><tbody>\'+list.map(x=>\'<tr><td>\'+esc(x.customer)+\'</td><td>\'+esc(money(x.amount))+\'</td><td>\'+esc(x.po)+\'<div class="secondary text-small">SO \'+esc(x.so||\'미정\')+\'</div></td><td>\'+esc(x.readyDate?x.readyDate.slice(5).replace(\'-\',\'/\'):\'미정\')+\'</td><td>\'+esc(x.shipFrom)+\'</td><td>\'+esc(x.status)+\'<div class="secondary text-small">\'+esc(x.note)+\'</div></td></tr>\').join(\'\')+(list.length?\'\':\'<tr><td colspan="6">해당 상태의 오더가 없어요.</td></tr>\')+\'</tbody></table></div>\':\'\';\n root.querySelectorAll(\'[data-order-status]\').forEach(b=>b.onclick=()=>{orderStatus=orderStatus===b.dataset.orderStatus?null:b.dataset.orderStatus;renderOrders();});\n}\nfunction sections(){return [...new Set([...(state.sections||[]),...state.items.map(x=>x.area||\'미분류\')])];}\nfunction visibleSections(){return sections().filter(x=>!(state.archivedSections||[]).includes(x));}\nfunction sectionColor(name){if(state.sectionColors?.[name])return state.sectionColors[name];const palette=[\'--blue\',\'--purple\',\'--green\',\'--orange\',\'--red\',\'--yellow\'];let hash=0;for(const c of name)hash=(hash*31+c.charCodeAt(0))>>>0;return palette[hash%palette.length];}\nfunction moveItem(id,area,before){const x=find(id);if(!x)return;const i=state.items.indexOf(x);state.items.splice(i,1);x.area=area;const target=before?state.items.findIndex(y=>y.id===before):-1;if(target>=0)state.items.splice(target,0,x);else{let last=-1;state.items.forEach((y,i)=>{if((y.area||\'미분류\')===area)last=i;});state.items.splice(last<0?state.items.length:last+1,0,x);}}\nconst esc=x=>String(x??\'\').replace(/[&<>"\']/g,c=>({\'&\':\'&amp;\',\'<\':\'&lt;\',\'>\':\'&gt;\',\'"\':\'&quot;\',"\'":\'&#39;\'}[c]));\nconst notify=s=>q(\'notice\').textContent=s;\nconst title=x=>[x.area,x.title].filter(Boolean).join(\' · \');\nconst safeLink=s=>{try{const u=new URL(s);return [\'https:\',\'http:\'].includes(u.protocol)?u.href:\'\';}catch{return \'\';}};\nconst bytes=x=>new TextEncoder().encode(JSON.stringify(x)).length;\nfunction event(text){state.events.unshift({id:crypto.randomUUID(),text,date:day()});state.events=state.events.slice(0,100);}\nfunction snapshot(){return structuredClone(state);}\nasync function mutate(fn,text){if(data.readOnly){notify(\'미리보기에서는 저장되지 않습니다.\');render();return false;}if(saving){notify(\'저장 중이에요. 잠시 후 다시 눌러주세요.\');return false;}const remote=read(bridge?.widgetState);if(valid(remote))state=structuredClone(remote);const before=structuredClone(state);fn();event(text);const payload={modelContent:{board:snapshot()}};if(!bridge?.setWidgetState){state=before;notify(\'공유 저장에 연결되지 않아 변경하지 않았어요.\');render();return false;}saving=true;render();try{await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));await bridge.setWidgetState(payload);notify(\'저장했어요\');return true;}catch(error){state=structuredClone(read(bridge.widgetState)||before);notify(error.message||\'저장 실패. 다시 시도해주세요.\');return false;}finally{saving=false;render();}}\nconst find=id=>state.items.find(x=>x.id===id);\nfunction sectionHistory(area){const list=state.items.filter(x=>(x.area||\'미분류\')===area&&[\'complete\',\'deleted\'].includes(x.status));return \'<details data-section-history="\'+esc(area)+\'"><summary class="cursor-interaction">완료 기록 (\'+list.length+\')</summary>\'+list.map(x=>\'<div class="line"><div class="grow">\'+esc(x.title)+\'<span class="secondary text-small"> · \'+\'완료\'+\'</span><div class="actions"><button class="btn" data-edit="\'+esc(x.id)+\'" type="button">노트 / 결과물</button><button class="btn" data-reactivate="\'+esc(x.id)+\'" type="button">다시 올리기</button></div></div></div>\').join(\'\')+(list.length?\'\':\'<div class="secondary">아직 기록이 없어요.</div>\')+\'</details>\';}\nfunction weekStart(){const d=new Date(viewDay+\'T12:00:00Z\');return plus(viewDay,-((d.getUTCDay()+6)%7));}\nfunction taskRow(x,area,showArea=false,inlineArea=false){return \'<tr data-item-row="\'+esc(x.id)+\'" data-item-area="\'+esc(area)+\'">\'+(showArea?\'<td><span class="deadline-area" style="border-left:4px solid var(\'+sectionColor(area)+\');padding-left:8px">\'+esc(area)+\'</span></td>\':\'\')+\'<td><div class="title-cell"><button class="btn btn-ghost" data-delete="\'+esc(x.id)+\'" type="button">완료</button><button class="btn btn-ghost" draggable="true" data-drag-item="\'+esc(x.id)+\'" type="button" aria-label="\'+esc(x.title)+\' 업무 이동">↕</button>\'+(inlineArea?\'<span style="border-left:4px solid var(\'+sectionColor(area)+\');padding-left:6px;white-space:nowrap">\'+esc(area)+\'</span>\':\'\')+\'<span class="task-name"><input class="form-control" data-title-id="\'+esc(x.id)+\'" value="\'+esc(x.title)+\'" maxlength="100" aria-label="업무 이름 수정"></span><label class="form-check"><input class="form-check-input" type="checkbox" data-week-id="\'+esc(x.id)+\'" \'+(x.weekGoal===weekStart()?\'checked\':\'\')+\'><span class="form-check-label">이번 주</span></label></div></td><td>\'+priorityCell(x,\'urgency\')+\'</td><td>\'+priorityCell(x,\'importance\')+\'</td><td><select class="form-select" data-owner-id="\'+esc(x.id)+\'" aria-label="담당자">\'+[\'미정\',\'Hugo\',\'제이나\',\'함께\'].map(a=>\'<option value="\'+esc(a)+\'" \'+(a===x.owner?\'selected\':\'\')+\'>\'+esc(personLabel(a))+\'</option>\').join(\'\')+\'</select></td><td>\'+dateCell(x,\'start\')+\'</td><td>\'+dateCell(x,\'due\')+\'</td><td><div class="manage-row">\'+buttons(x)+\'<span class="section-move"><select class="form-select" aria-label="섹션 이동" data-area-id="\'+esc(x.id)+\'">\'+visibleSections().map(a=>\'<option \'+(a===area?\'selected\':\'\')+\'>\'+esc(a)+\'</option>\').join(\'\')+\'</select></span></div></td></tr>\';}\nfunction dateCell(x,key){const date=x[key]||\'\';return \'<div class="date-edit"><button type="button" class="btn btn-ghost" data-date-open="\'+esc(x.id)+\'" data-date-key="\'+key+\'" aria-label="\'+esc(x.title)+\' \'+(key===\'start\'?\'시작일\':\'마감일\')+\' 더블클릭해서 수정">\'+esc(date?date.slice(5).replace(\'-\',\'/\'):\'미정\')+\'</button><input hidden type="date" class="form-control" data-date="\'+key+\'" data-date-id="\'+esc(x.id)+\'" value="\'+esc(date)+\'" aria-label="\'+esc(x.title)+\' \'+(key===\'start\'?\'시작일\':\'마감일\')+\'"></div>\';}\nfunction priorityCell(x,key){const options=key===\'urgency\'?[\'미정\',\'낮음\',\'보통\',\'높음\',\'긴급\']:[\'미정\',\'낮음\',\'보통\',\'높음\',\'매우 높음\'];const value=x[key]||\'미정\';const color={\'미정\':\'#999\',\'낮음\':\'#43c67b\',\'보통\':\'#e6c94c\',\'높음\':\'#ff6868\',\'긴급\':\'#ff6868\',\'매우 높음\':\'#ff6868\'};return \'<select class="form-select priority-dot" style="color:\'+color[value]+\'" data-priority="\'+key+\'" data-priority-id="\'+esc(x.id)+\'" aria-label="\'+esc(x.title)+\' \'+(key===\'urgency\'?\'긴급도\':\'중요도\')+\'">\'+options.map(a=>\'<option value="\'+a+\'" style="color:\'+color[a]+\'" \'+(a===value?\'selected\':\'\')+\'>● \'+a+\'</option>\').join(\'\')+\'</select>\';}\n\nfunction buttons(x){return \'<div class="actions"><button class="btn" data-edit="\'+esc(x.id)+\'" type="button">노트 / 결과물\'+(x.note||x.result||x.link?\' •\':\'\')+\'</button></div>\';}\nfunction short(x){return \'<div class="line"><button class="btn btn-ghost" \'+(x.status===\'complete\'?\'data-reactivate\':\'data-delete\')+\'="\'+esc(x.id)+\'" type="button">\'+(x.status===\'complete\'?\'다시 올리기\':\'완료\')+\'</button><div class="grow"><strong>\'+esc(title(x))+\'</strong><div class="secondary text-small">담당 \'+esc(personLabel(x.owner))+\' · \'+(x.doneDate?\'완료 \'+esc(x.doneDate):\'등록 \'+esc(x.added))+\'</div>\'+buttons(x)+\'</div></div>\';}\nfunction openEdit(id){if(editId===id&&!editorNode.hidden){if(dirtyEditor)noteDrafts.set(id,{_base:editBase,owner:q(\'edit-owner\').value,start:q(\'edit-start\').value,due:q(\'edit-due\').value,note:q(\'edit-note\').value,result:q(\'edit-result\').value,link:q(\'edit-link\').value,emailSubject:q(\'edit-email\').value});editorNode.hidden=true;root.append(editorNode);root.querySelectorAll(\'[data-inline-editor]\').forEach(x=>x.remove());editId=null;dirtyEditor=false;return;}if(dirtyEditor){if(editId!==id)notify(\'작성 중인 노트를 저장하거나 닫은 뒤 다른 업무를 열어주세요.\');return;}const original=find(id);if(!original)return;editBase=structuredClone(noteDrafts.get(id)?._base||original);const x={...original,...(noteDrafts.get(id)||{})};root.append(editorNode);root.querySelectorAll(\'[data-inline-editor]\').forEach(x=>x.remove());editId=id;dirtyEditor=noteDrafts.has(id);editorNode.hidden=false;q(\'edit-title\').textContent=\'노트 / 결과물\';q(\'edit-owner\').value=x.owner;q(\'edit-start\').value=x.start||\'\';q(\'edit-due\').value=x.due||\'\';q(\'edit-note\').value=x.note||\'\';q(\'edit-result\').value=x.result||\'\';q(\'edit-link\').value=x.link||\'\';q(\'edit-email\').value=x.emailSubject||\'\';const link=safeLink(x.link);q(\'result-preview\').innerHTML=link?\'<a href="\'+esc(link)+\'" target="_blank" rel="noopener noreferrer">결과물 열기 ↗</a>\':\'\';mountEditor();q(\'edit-note\').focus({preventScroll:true});}\nfunction render(){renderOrders();q(\'my-profile\').textContent=personLabel(actor.display_name);q(\'my-emoji\').innerHTML=emojiChoices.map(e=>\'<option value="\'+e+\'" \'+(e===(state.memberEmoji?.[actor.display_name]||\'\')?\'selected\':\'\')+\'>\'+(e||\'없음\')+\'</option>\').join(\'\');q(\'my-emoji\').disabled=saving||data.readOnly;for(const id of [\'task-owner\',\'edit-owner\'])for(const option of q(id).options){const name=option.getAttribute(\'value\')||option.textContent;option.value=name;option.textContent=personLabel(name);}addHome.append(addForm);root.querySelectorAll(\'[data-inline-add]\').forEach(x=>x.remove());root.append(editorNode);root.querySelectorAll(\'[data-inline-editor]\').forEach(x=>x.remove());q(\'date-label\').textContent=viewDay+(viewDay===day()?\' · 오늘\':\' · 미리보기\');q(\'real-day\').hidden=viewDay===day();const selectedAddedDay=addedDay||day();q(\'new-title\').textContent=selectedAddedDay.slice(5).replace(\'-\',\'/\')+(selectedAddedDay===day()?\' 오늘 새로 추가된 일\':\' 추가된 일\');const dates=[...new Set([day(),...Array.from({length:7},(_,i)=>plus(day(),-i)),...state.items.filter(x=>!x.imported).map(x=>x.added)])].filter(Boolean).sort((a,b)=>b.localeCompare(a));q(\'added-dates\').innerHTML=dates.map(d=>\'<button class="btn btn-ghost" type="button" data-added-day="\'+esc(d)+\'" aria-pressed="\'+(d===selectedAddedDay)+\'">\'+esc(d.slice(5).replace(\'-\',\'/\'))+(d===day()?\' 오늘\':\'\')+\'</button>\').join(\'\');root.querySelectorAll(\'[data-added-day]\').forEach(b=>b.onclick=()=>{addedDay=b.dataset.addedDay;render();});\nconst chosen=q(\'task-section\').value; q(\'task-section\').innerHTML=visibleSections().map(a=>\'<option value="\'+esc(a)+\'">\'+esc(a)+\'</option>\').join(\'\')+\'<option value="__new__">＋ 새 섹션 만들기</option>\';q(\'task-section\').value=(chosen===\'__new__\'||visibleSections().includes(chosen))?chosen:\'미분류\';q(\'new-section-label\').hidden=q(\'task-section\').value!==\'__new__\';\nconst fresh=state.items.filter(x=>!completions.has(x.id)&&!x.imported&&x.added===selectedAddedDay);q(\'new-list\').innerHTML=fresh.map(short).join(\'\')||\'<div class="secondary">이 날짜에 새로 추가한 일이 없어요.</div>\';\nconst daily=state.items.filter(x=>x.daily&&x.added<=viewDay&&x.status===\'active\'&&!completions.has(x.id));q(\'daily-list\').innerHTML=\'<div class="daily-grid daily-header"><span class="secondary text-small">내 계정으로 체크하면 내 이름과 색깔로 표시됩니다.</span>\'+people.map(p=>\'<span>\'+esc(personLabel(p))+\'</span>\').join(\'\')+\'<span></span></div>\'+daily.map(x=>\'<div class="daily-grid"><span><button class="btn btn-ghost" data-delete="\'+esc(x.id)+\'" type="button">완료</button> \'+esc(x.title)+\'</span>\'+people.map(person=>\'<label class="progress-check"><input class="cursor-interaction" type="checkbox" data-daily="\'+esc(x.id)+\'" data-person="\'+person+\'" aria-label="\'+esc(x.title)+\' \'+person+\' 진행 중" \'+(x.checks?.[viewDay]?.people?.[person]?\'checked\':\'\')+\'></label>\').join(\'\')+\'<button class="btn btn-ghost" data-edit="\'+esc(x.id)+\'" type="button">노트 / 결과물</button><div class="inline-slot grow"></div></div>\').join(\'\');\nconst active=state.items.filter(x=>!x.daily&&x.status===\'active\'&&!completions.has(x.id)&&(x.imported||x.added<=viewDay));q(\'active-title\').textContent=\'전체 업무표 · \'+active.length;\nconst focus=state.items.filter(x=>!x.daily&&x.status===\'active\'&&!completions.has(x.id)&&visibleSections().includes(x.area||\'미분류\')&&([\'높음\',\'긴급\'].includes(x.urgency)||(x.due&&x.due<=plus(weekStart(),6))||x.weekGoal===weekStart()));q(\'focus-list\').innerHTML=focus.map(x=>taskRow(x,x.area||\'미분류\',false,true)).join(\'\')||\'<tr><td colspan="7" class="secondary">업무의 ‘이번 주’를 체크하거나 긴급도·마감일을 지정해주세요.</td></tr>\';\nconst deadlines=state.items.filter(x=>!x.daily&&x.status===\'active\'&&!completions.has(x.id)&&x.due&&visibleSections().includes(x.area||\'미분류\')).sort((a,b)=>a.due.localeCompare(b.due)||a.title.localeCompare(b.title));q(\'deadline-list\').innerHTML=deadlines.map(x=>taskRow(x,x.area||\'미분류\',true)).join(\'\')||\'<tr><td colspan="8" class="secondary">마감일을 지정한 업무가 여기에 표시됩니다.</td></tr>\';\nq(\'active-list\').innerHTML=visibleSections().map(area=>\'<tr data-section="\'+esc(area)+\'"><td colspan="7"><div class="viz-row"><h2 style="border-left:6px solid var(\'+sectionColor(area)+\');padding-left:12px">\'+esc(area)+\'</h2><button class="btn" data-add-section="\'+esc(area)+\'" type="button">＋ 새 업무</button><button class="btn btn-ghost" draggable="true" data-drag-section="\'+esc(area)+\'" type="button" aria-label="\'+esc(area)+\' 섹션 순서 이동">↕</button><span class="secondary text-small">\'+active.filter(x=>(x.area||\'미분류\')===area).length+\'개</span><button class="btn btn-ghost" data-section-up="\'+esc(area)+\'" type="button" aria-label="\'+esc(area)+\' 섹션 위로">↑</button><button class="btn btn-ghost" data-section-down="\'+esc(area)+\'" type="button" aria-label="\'+esc(area)+\' 섹션 아래로">↓</button><button class="btn" data-rename="\'+esc(area)+\'" type="button">이름 수정</button><button class="btn btn-ghost" data-archive-section="\'+esc(area)+\'" type="button">Archive</button>\'+sectionHistory(area)+\'</div><form data-rename-form="\'+esc(area)+\'" class="viz-row" hidden><label class="form-label">섹션 이름<input class="form-control" maxlength="40" required value="\'+esc(area)+\'"></label><button class="btn" type="submit">저장</button></form></td></tr><tr class="section-column-labels"><th>업무</th><th>긴급도</th><th>중요도</th><th>담당</th><th>시작일</th><th>마감</th><th>관리</th></tr>\'+active.filter(x=>(x.area||\'미분류\')===area).map(x=>taskRow(x,area)).join(\'\')).join(\'\');\nconst complete=state.items.filter(x=>x.status===\'complete\');q(\'complete-title\').textContent=\'완료함 · \'+complete.length;q(\'complete-list\').innerHTML=complete.map(short).join(\'\')||\'<div class="secondary">완료한 업무가 여기에 모여요.</div>\';\nq(\'deleted-list\').innerHTML=\'\';\nq(\'event-list\').innerHTML=state.events.map(x=>\'<div class="line"><div>\'+esc(x.text)+\'<div class="secondary text-small">\'+esc(x.date)+\' · \'+esc(x.actor||\'이전 기록\')+\'</div></div></div>\').join(\'\')||\'<div class="secondary">새로 추가하거나 완료하면 기록이 남아요.</div>\';q(\'news-toggle\').textContent=\'🔔 새 소식 \'+state.events.length;\nq(\'daily-history\').innerHTML=state.items.filter(x=>x.daily).flatMap(x=>Object.entries(x.checks||{}).sort(([a],[b])=>b.localeCompare(a)).map(([d,c])=>\'<div class="line">\'+esc(d)+\' · \'+esc(x.title)+\' · \'+esc(c.people?Object.keys(c.people).join(\', \')||\'진행 표시 해제\':\'이전 체크 기록\')+\'</div>\')).join(\'\')||\'<div class="secondary">Daily 진행 기록이 없어요.</div>\';\nroot.querySelectorAll(\'[data-edit]\').forEach(b=>b.onclick=()=>openEdit(b.dataset.edit));\nroot.querySelectorAll(\'[data-reactivate]\').forEach(b=>b.onclick=()=>mutate(()=>{const item=find(b.dataset.reactivate);item.status=\'active\';item.doneDate=\'\';},find(b.dataset.reactivate).title+\' 다시 올렸어요\'));\nconst pending=state.items.filter(x=>x.needsReview&&x.status!==\'deleted\');q(\'review-panel\').hidden=!pending.length;q(\'review-title\').textContent=\'새 노트 · 결과물 확인 필요 · \'+pending.length;\nq(\'review-list\').innerHTML=pending.map(x=>\'<div class="line"><div class="grow"><strong>\'+esc(title(x))+\'</strong><span class="secondary text-small"> · \'+esc(x.noteUpdatedAt||\'\')+\'</span><div class="actions"><button class="btn" data-review-open="\'+esc(x.id)+\'" type="button">내용 보기</button><button class="btn" data-reviewed="\'+esc(x.id)+\'" type="button">확인했어요</button></div></div></div>\').join(\'\');\nroot.querySelectorAll(\'[data-review-open]\').forEach(b=>b.onclick=()=>{const target=root.querySelector(\'[data-item-row="\'+b.dataset.reviewOpen+\'"] [data-edit]\');if(target)openEdit(b.dataset.reviewOpen);else{b.dataset.edit=b.dataset.reviewOpen;openEdit(b.dataset.reviewOpen);}});\nroot.querySelectorAll(\'[data-reviewed]\').forEach(b=>b.onclick=()=>mutate(()=>{find(b.dataset.reviewed).needsReview=false;find(b.dataset.reviewed).reviewedAt=day();},find(b.dataset.reviewed).title+\' 새 노트 / 결과물을 확인했어요\'));\nq(\'archived-sections\').innerHTML=(state.archivedSections||[]).map(area=>\'<div class="line"><div class="grow">\'+esc(area)+\' · \'+state.items.filter(x=>(x.area||\'미분류\')===area&&x.status===\'active\'&&!completions.has(x.id)).length+\'개</div><button class="btn" data-unarchive="\'+esc(area)+\'" type="button">꺼내기</button></div>\').join(\'\')||\'<div class="secondary">보관한 섹션이 없어요.</div>\';\nroot.querySelectorAll(\'[data-archive-section]\').forEach(b=>b.onclick=()=>mutate(()=>{state.archivedSections=[...new Set([...(state.archivedSections||[]),b.dataset.archiveSection])];},b.dataset.archiveSection+\' 섹션을 보관했어요\'));\nroot.querySelectorAll(\'[data-unarchive]\').forEach(b=>b.onclick=()=>mutate(()=>{state.archivedSections=(state.archivedSections||[]).filter(x=>x!==b.dataset.unarchive);},b.dataset.unarchive+\' 섹션을 꺼냈어요\'));\nroot.querySelectorAll(\'[data-status]\').forEach(b=>b.onclick=async()=>{const id=b.dataset.status;const x=find(id);await mutate(()=>{const item=find(id);item.status=item.status===\'complete\'?\'active\':\'complete\';item.doneDate=item.status===\'complete\'?day():\'\';},title(x)+(x.status===\'complete\'?\' 다시 열었어요\':\' 완료했어요\'));});\nroot.querySelectorAll(\'[data-delete]\').forEach(b=>b.onclick=()=>completeNow(b.dataset.delete));\nroot.querySelectorAll(\'[data-restore]\').forEach(b=>b.onclick=async()=>{const id=b.dataset.restore,x=find(id);await mutate(()=>{const item=find(id);item.status=item.previousStatus||\'active\';},title(x)+\' 복원했어요\');});\nroot.querySelectorAll(\'[data-daily]\').forEach(b=>b.onchange=async()=>{const id=b.dataset.daily,on=b.checked,person=b.dataset.person,x=find(id);await mutate(()=>{const item=find(id);item.checks=item.checks||{};const c=item.checks[viewDay]||{date:viewDay};c.people=c.people||{};if(on)c.people[person]=true;else delete c.people[person];item.checks[viewDay]=c;},title(x)+\' \'+person+(on?\' 진행 중으로 표시했어요\':\' 진행 표시를 해제했어요\'));});\nroot.querySelectorAll(\'[data-date]\').forEach(b=>b.onchange=async()=>{const value=b.value,id=b.dataset.dateId,key=b.dataset.date;await mutate(()=>{find(id)[key]=value;},find(id).title+\' \'+(key===\'start\'?\'시작일\':\'마감일\')+\'을 바꿨어요\');});\nroot.querySelectorAll(\'[data-date-open]\').forEach(b=>{b.ondblclick=()=>{const input=b.parentElement.querySelector(\'input\');b.hidden=true;input.hidden=false;input.focus();try{input.showPicker();}catch{}};});\nroot.querySelectorAll(\'[data-week-id]\').forEach(b=>b.onchange=async()=>{const id=b.dataset.weekId,on=b.checked;await mutate(()=>{find(id).weekGoal=on?weekStart():\'\';},find(id).title+(on?\' 이번 주 목표로 표시했어요\':\' 이번 주 목표를 해제했어요\'));});\nroot.querySelectorAll(\'[data-priority]\').forEach(b=>b.onchange=async()=>{const id=b.dataset.priorityId,key=b.dataset.priority,value=b.value;await mutate(()=>{find(id)[key]=value;},find(id).title+\' \'+(key===\'urgency\'?\'긴급도\':\'중요도\')+\'를 \'+value+\'로 바꿨어요\');});\nroot.querySelectorAll(\'[data-title-id]\').forEach(b=>b.onchange=async()=>{const value=b.value.trim(),id=b.dataset.titleId;if(!value){b.value=find(id).title;return;}await mutate(()=>{find(id).title=value;},\'업무 이름을 \'+value+\'으로 바꿨어요\');});\nroot.querySelectorAll(\'[data-owner-id]\').forEach(b=>b.onchange=async()=>{const value=b.value,id=b.dataset.ownerId;await mutate(()=>{find(id).owner=value;},find(id).title+\' 담당을 \'+value+\'로 바꿨어요\');});\nroot.querySelectorAll(\'[data-area-id]\').forEach(b=>b.onchange=async()=>{const area=b.value,id=b.dataset.areaId;await mutate(()=>moveItem(id,area),find(id).title+\'을 \'+area+\' 섹션으로 옮겼어요\');});\nroot.querySelectorAll(\'[data-add-section]\').forEach(b=>b.onclick=()=>{const area=b.dataset.addSection;addSection=addSection===area?null:area;q(\'task-section\').value=area;q(\'new-section-label\').hidden=true;q(\'task-new-section\').required=false;q(\'task-daily\').checked=false;mountAdd();q(\'add-close\').hidden=!addSection;if(addSection)q(\'task-title\').focus({preventScroll:true});});\nroot.querySelectorAll(\'[data-rename]\').forEach(b=>b.onclick=()=>{const form=Array.from(root.querySelectorAll(\'[data-rename-form]\')).find(x=>x.dataset.renameForm===b.dataset.rename);form.hidden=!form.hidden;if(!form.hidden)form.querySelector(\'input\').focus();});\nroot.querySelectorAll(\'[data-rename-form]\').forEach(form=>form.onsubmit=async e=>{e.preventDefault();const old=form.dataset.renameForm,name=form.querySelector(\'input\').value.trim();if(!name||name===old){form.hidden=true;return;}if(sections().includes(name)){notify(\'같은 이름의 섹션이 있어요.\');return;}await mutate(()=>{const color=sectionColor(old);state.sectionColors=state.sectionColors||{};state.sectionColors[name]=color;delete state.sectionColors[old];state.sections=sections().map(x=>x===old?name:x);state.items.forEach(x=>{if((x.area||\'미분류\')===old)x.area=name;});},old+\' 섹션 이름을 \'+name+\'으로 바꿨어요\');});\nroot.querySelectorAll(\'[data-drag-item],[data-drag-section]\').forEach(b=>{b.ondragstart=e=>{if(dirtyEditor||saving){e.preventDefault();notify(\'작성 중인 노트를 먼저 저장해주세요.\');return;}drag=b.dataset.dragItem?{type:\'item\',id:b.dataset.dragItem}:{type:\'section\',area:b.dataset.dragSection};e.dataTransfer.setData(\'text/plain\',JSON.stringify(drag));e.dataTransfer.effectAllowed=\'move\';};b.ondragend=()=>drag=null;});\nroot.querySelectorAll(\'[data-section],[data-item-row]\').forEach(b=>{b.ondragover=e=>{if(drag){e.preventDefault();e.dataTransfer.dropEffect=\'move\';}};b.ondrop=async e=>{e.preventDefault();if(!drag)return;const d=drag;drag=null;const area=b.dataset.section||b.dataset.itemArea;if(d.type===\'section\'){await mutate(()=>{const list=sections();const from=list.indexOf(d.area),to=list.indexOf(area);if(from<0||to<0||from===to)return;list.splice(from,1);list.splice(list.indexOf(area),0,d.area);state.sections=list;},\'섹션 순서를 바꿨어요\');}else if(d.id!==b.dataset.itemRow)await mutate(()=>moveItem(d.id,area,b.dataset.itemRow),\'업무 위치를 바꿨어요\');};});\nfor(const [attr,delta] of [[\'sectionUp\',-1],[\'sectionDown\',1]])root.querySelectorAll(\'[data-\'+(delta<0?\'section-up\':\'section-down\')+\']\').forEach(b=>b.onclick=()=>mutate(()=>{const list=sections(),i=list.indexOf(b.dataset[attr]),j=i+delta;if(j>=0&&j<list.length){[list[i],list[j]]=[list[j],list[i]];state.sections=list;}},\'섹션 순서를 바꿨어요\'));\nfor(const [attr,delta] of [[\'itemUp\',-1],[\'itemDown\',1]])root.querySelectorAll(\'[data-\'+(delta<0?\'item-up\':\'item-down\')+\']\').forEach(b=>b.onclick=()=>mutate(()=>{const id=b.dataset[attr],x=find(id),list=state.items.filter(y=>y.status===\'active\'&&!y.daily&&(y.area||\'미분류\')===(x.area||\'미분류\')),i=list.findIndex(y=>y.id===id),other=list[i+delta];if(other){const a=state.items.indexOf(x),z=state.items.indexOf(other);[state.items[a],state.items[z]]=[state.items[z],state.items[a]];}},\'업무 순서를 바꿨어요\'));\nroot.style.setProperty(\'--daily-columns\',people.length);root.querySelectorAll(\'[data-daily]\').forEach(b=>{b.disabled=saving||b.dataset.person!==actor.display_name||data.readOnly;b.style.accentColor=data.members.find(m=>m.display_name===b.dataset.person)?.color||\'#4285f4\';});root.querySelectorAll(\'button\').forEach(b=>{if(b.id!==\'news-toggle\'&&!b.hasAttribute(\'data-delete\'))b.disabled=saving;});\nmountEditor();mountAdd();\n}\nq(\'my-emoji\').onchange=async()=>{const emoji=q(\'my-emoji\').value;if(!emojiChoices.includes(emoji))return;await mutate(()=>{state.memberEmoji=state.memberEmoji||{};state.memberEmoji[actor.display_name]=emoji;},actor.display_name+\' 이모지를 바꿨어요\');};\nq(\'new-toggle\').onclick=()=>{const content=q(\'new-content\');content.hidden=!content.hidden;q(\'new-toggle\').textContent=content.hidden?\'▸ 펼치기\':\'▾ 접기\';q(\'new-toggle\').setAttribute(\'aria-expanded\',String(!content.hidden));};\nq(\'daily-toggle\').onclick=()=>{const content=q(\'daily-content\');content.hidden=!content.hidden;q(\'daily-toggle\').textContent=content.hidden?\'▸ 펼치기\':\'▾ 접기\';q(\'daily-toggle\').setAttribute(\'aria-expanded\',String(!content.hidden));};\nq(\'daily-add-toggle\').onclick=()=>{q(\'daily-content\').hidden=false;q(\'daily-toggle\').textContent=\'▾ 접기\';q(\'daily-toggle\').setAttribute(\'aria-expanded\',\'true\');const form=q(\'daily-add-form\');form.hidden=!form.hidden;q(\'daily-add-toggle\').setAttribute(\'aria-expanded\',String(!form.hidden));if(!form.hidden)q(\'daily-add-title\').focus({preventScroll:true});};\nq(\'daily-add-close\').onclick=()=>{q(\'daily-add-form\').hidden=true;q(\'daily-add-toggle\').setAttribute(\'aria-expanded\',\'false\');};\nq(\'daily-add-form\').onsubmit=async e=>{e.preventDefault();const t=q(\'daily-add-title\').value.trim();if(!t)return;const input={id:crypto.randomUUID(),area:\'미분류\',title:t,owner:\'미정\',start:day(),due:\'\',added:day(),note:\'\',result:\'\',link:\'\',daily:true,checks:{},status:\'active\'};const ok=await mutate(()=>state.items.push(input),t+\' Daily 할 일에 추가했어요\');if(ok){q(\'daily-add-title\').value=\'\';q(\'daily-add-form\').hidden=true;q(\'daily-add-toggle\').setAttribute(\'aria-expanded\',\'false\');viewDay=day();render();}};\nq(\'add-form\').onsubmit=async e=>{e.preventDefault();const t=q(\'task-title\').value.trim();if(!t)return;const create=q(\'task-section\').value===\'__new__\',area=create?q(\'task-new-section\').value.trim():q(\'task-section\').value;if(!area){notify(\'새 섹션 이름을 입력해주세요.\');return;}if(create&&sections().includes(area)){notify(\'이미 있는 섹션이에요. 목록에서 선택해주세요.\');return;}if(create&&sections().length>=25){notify(\'섹션은 25개까지 만들 수 있어요.\');return;}const input={id:crypto.randomUUID(),area,title:t,owner:q(\'task-owner\').value,start:day(),due:q(\'task-due\').value,added:day(),note:\'\',result:\'\',link:\'\',daily:q(\'task-daily\').checked,checks:{},status:\'active\'};const ok=await mutate(()=>{if(create)state.sections=[...sections(),area];state.items.push(input);},t+\' 새로 추가했어요\');if(ok){q(\'task-title\').value=\'\';viewDay=day();addedDay=null;addSection=null;q(\'add-close\').hidden=true;render();}};\nq(\'add-close\').onclick=()=>{addSection=null;mountAdd();q(\'add-close\').hidden=true;};addHome.ontoggle=()=>{if(addHome.open&&addSection){addSection=null;mountAdd();q(\'add-close\').hidden=true;}};\nq(\'task-section\').onchange=()=>{q(\'new-section-label\').hidden=q(\'task-section\').value!==\'__new__\';q(\'task-new-section\').required=q(\'task-section\').value===\'__new__\';};\nq(\'edit-form\').oninput=()=>dirtyEditor=true;\nq(\'edit-form\').onsubmit=async e=>{e.preventDefault();const id=editId,x=find(id);if(!x)return;const link=q(\'edit-link\').value.trim();if(link&&!safeLink(link)){notify(\'결과물 링크는 http 또는 https 주소로 적어주세요.\');return;}const values={owner:q(\'edit-owner\').value,start:q(\'edit-start\').value,due:q(\'edit-due\').value,note:q(\'edit-note\').value,result:q(\'edit-result\').value,emailSubject:q(\'edit-email\').value,link};const remoteItem=read(bridge.widgetState)?.items.find(item=>item.id===id);const updates=Object.fromEntries(Object.entries(values).filter(([k,v])=>v!==(editBase?.[k]||\'\')));if(Object.keys(updates).some(k=>(remoteItem?.[k]||\'\')!==(editBase?.[k]||\'\'))){notify(\'다른 팀원이 이 노트를 수정했어요. 작성한 내용은 유지됩니다. 최신 내용을 확인하고 다시 작성해주세요.\');return;}const changed=[\'note\',\'result\',\'link\',\'emailSubject\'].some(k=>(x[k]||\'\')!==(values[k]||\'\'));const ok=await mutate(()=>{Object.assign(find(id),updates);if(changed){find(id).needsReview=true;find(id).noteUpdatedAt=day();}},title(x)+\' 노트 / 결과물을 업데이트했어요\');if(ok){noteDrafts.delete(id);dirtyEditor=false;editId=null;openEdit(id);}};\nq(\'edit-close\').onclick=()=>{editorNode.hidden=true;root.append(editorNode);root.querySelectorAll(\'[data-inline-editor]\').forEach(x=>x.remove());editId=null;dirtyEditor=false;};\nq(\'news-toggle\').onclick=()=>q(\'news-panel\').hidden=!q(\'news-panel\').hidden;q(\'next-day\').onclick=()=>{viewDay=plus(viewDay,1);render();};q(\'real-day\').onclick=()=>{viewDay=day();render();};\nparentElement._operation={root,receive(nextData){\n  bridge.widgetState={modelContent:{board:nextData.board}};\n  const changed=revision!==nextData.revision;revision=nextData.revision;\n  if(pending&&nextData.ack?.id===pending.id){const request=pending;pending=null;if(nextData.ack.ok)request.resolve();else request.reject(new Error(nextData.ack.message||\'다른 팀원이 먼저 수정했어요. 최신 내용으로 다시 시도해주세요.\'));return;}\n  const next=read(bridge.widgetState);if(!valid(next)||saving)return;\n  const focused=parentElement.activeElement;\n  if(changed&&!dirtyEditor&&!(focused&&[\'INPUT\',\'TEXTAREA\',\'SELECT\'].includes(focused.tagName))){const fresh=next.events.filter(x=>!seen.has(x.id));state=structuredClone(next);seen=new Set(next.events.map(x=>x.id));render();if(fresh.length)notify((fresh[0].actor?fresh[0].actor+\' · \':\'\')+fresh[0].text);}\n}};\nlet last=day();\nconst poll=setInterval(()=>{if(!root.isConnected){clearInterval(poll);return;}const now=day();if(now!==last){if(viewDay===last)viewDay=now;last=now;render();}if(!saving&&!dirtyEditor)setTriggerValue(\'refresh\',Date.now());},5000);\nrender();\n}\n'

@st.cache_resource
def login_storage():
    return st.components.v2.component("operation_login_storage", html="<span></span>", js="""
export default function({parentElement,data,setTriggerValue}) {
  if (parentElement._loginCommand === data.id) return;
  parentElement._loginCommand = data.id;
  try {
    let token = null;
    if (data.action === 'store') localStorage.setItem(data.key, data.token);
    if (data.action === 'clear') localStorage.removeItem(data.key);
    if (data.action === 'restore') token = localStorage.getItem(data.key);
    setTriggerValue('reply', {id:data.id, token, ok:true});
  } catch (_) { setTriggerValue('reply', {id:data.id, ok:false}); }
}
""")


def restore_login():
    if "login_storage_command" not in st.session_state:
        st.session_state.login_storage_command = {"id": str(uuid.uuid4()), "action": "restore"}
    command = st.session_state.login_storage_command
    result = login_storage()(data={**command, "key": "operation-login:" + URL}, key="login_storage", on_reply_change=lambda: None, height=0)
    reply = result.get("reply")
    if reply and reply.get("id") == command["id"] and st.session_state.get("login_storage_handled") != command["id"]:
        st.session_state.login_storage_handled = command["id"]
        st.session_state.login_storage_ready = True
        if command["action"] == "restore" and reply.get("token"):
            try:
                session = api("/auth/v1/token?grant_type=refresh_token", "POST", {"refresh_token": reply["token"]})
                api("/auth/v1/user", token=session["access_token"])
                session.setdefault("expires_at", time.time() + session.get("expires_in", 3600))
                st.session_state.auth = session
                st.session_state.remember_login = True
                queue_login_storage("store", session["refresh_token"])
                st.rerun()
            except (RuntimeError, urllib.error.URLError):
                queue_login_storage("clear")
                st.rerun()
        if not reply.get("ok"):
            st.caption("브라우저의 저장 제한으로 로그인 유지가 안 될 수 있습니다.")
    if not st.session_state.get("login_storage_ready"):
        st.caption("로그인 상태를 확인하고 있어요…")
        if st.button("로그인 화면으로 계속"):
            st.session_state.login_storage_ready = True
            st.rerun()
        st.stop()


def queue_login_storage(action, token=None):
    st.session_state.login_storage_command = {"id": str(uuid.uuid4()), "action": action, "token": token}


@st.cache_resource
def board_component(html, js, css):
    version = hashlib.sha256((html + js + css).encode()).hexdigest()[:12]
    return st.components.v2.component(
        "operation_board_" + version, html=html, js=js, css=css,
    )


preview = False
if not URL or not KEY:
    st.error("서버 연결 설정이 필요합니다.")
    st.stop()
st.title("Operation")
if not preview:
    restore_login()
if not preview and not st.session_state.get("auth"):
    st.caption("회사 이메일로 로그인하면 체크와 수정 기록에 내 이름이 남습니다.")
    with st.form("login"):
        email = st.text_input("이메일")
        password = st.text_input("비밀번호", type="password")
        remember_login = st.checkbox("이 브라우저에서 로그인 유지", value=True)
        login = st.form_submit_button("로그인", type="primary")
    if login:
        try:
            session = api("/auth/v1/token?grant_type=password", "POST", {
                "email": email.strip(), "password": password,
            })
            session.setdefault("expires_at", time.time() + session.get("expires_in", 3600))
            st.session_state.auth = session
            st.session_state.remember_login = remember_login
            queue_login_storage("store" if remember_login else "clear", session.get("refresh_token") if remember_login else None)
            st.rerun()
        except RuntimeError:
            st.error("로그인하지 못했어요. 이메일·비밀번호와 이메일 인증 여부를 확인해주세요.")
    with st.expander("처음 사용하는 경우 · 계정 만들기"):
        with st.form("signup"):
            signup_email = st.text_input("회사 이메일", key="signup_email")
            signup_password = st.text_input("비밀번호 (8자 이상)", type="password", key="signup_password")
            signup = st.form_submit_button("계정 만들기")
        if signup:
            if len(signup_password) < 8:
                st.error("비밀번호를 8자 이상 입력해주세요.")
            else:
                try:
                    api("/auth/v1/signup", "POST", {"email": signup_email.strip(), "password": signup_password})
                    st.info("신규 계정이면 인증 메일을 확인해주세요. 이미 가입한 이메일은 다시 가입해도 비밀번호가 바뀌지 않습니다. 비밀번호를 모르면 아래 ‘비밀번호 재설정’을 이용해주세요.")
                except RuntimeError:
                    st.error("가입 요청을 처리하지 못했어요. Supabase 이메일 가입 설정을 확인해주세요.")
    with st.expander("비밀번호 재설정"):
        st.caption("회사 이메일로 인증번호를 받은 뒤 새 비밀번호를 정해주세요.")
        with st.form("recovery_send"):
            recovery_email = st.text_input("재설정할 이메일").strip().lower()
            send_recovery = st.form_submit_button("인증번호 이메일 보내기")
        if send_recovery:
            if not recovery_email or "@" not in recovery_email:
                st.error("이메일을 입력해주세요.")
            elif time.time() < st.session_state.get("recovery_sent_at", 0) + 60:
                st.warning("재발송은 1분 뒤에 시도해주세요.")
            else:
                try:
                    api("/auth/v1/recover", "POST", {"email": recovery_email})
                    st.session_state.recovery_email = recovery_email
                    st.session_state.recovery_sent_at = time.time()
                    st.session_state.pop("recovery_session", None)
                    st.success("등록된 이메일이면 재설정 메일이 발송됩니다. 스팸함도 확인해주세요.")
                except (RuntimeError, urllib.error.URLError) as error:
                    st.error("메일 발송 요청 실패: " + str(error))
        with st.form("recovery_verify", clear_on_submit=True):
            verify_email = st.text_input("인증할 이메일", value=st.session_state.get("recovery_email", "")).strip().lower()
            recovery_code = st.text_input("메일의 인증번호", type="password")
            new_password = st.text_input("새 비밀번호 (8자 이상)", type="password")
            confirm_password = st.text_input("새 비밀번호 확인", type="password")
            reset_password = st.form_submit_button("비밀번호 변경")
        if reset_password:
            if len(new_password) < 8 or new_password != confirm_password:
                st.error("8자 이상의 같은 비밀번호를 두 번 입력해주세요.")
            else:
                try:
                    recovery = st.session_state.get("recovery_session", {})
                    if recovery.get("email") != verify_email or recovery.get("expires_at", 0) < time.time():
                        verified = api("/auth/v1/verify", "POST", {"email": verify_email, "token": recovery_code.strip(), "type": "recovery"})
                        recovery = {"email": verify_email, "access_token": verified["access_token"], "expires_at": time.time() + min(verified.get("expires_in", 3600), 600)}
                        st.session_state.recovery_session = recovery
                    api("/auth/v1/user", "PUT", {"password": new_password}, token=recovery["access_token"])
                    st.session_state.pop("recovery_session", None)
                    st.success("비밀번호를 변경했어요. 위 로그인 칸에서 새 비밀번호로 로그인해주세요.")
                except (RuntimeError, urllib.error.URLError) as error:
                    st.error("비밀번호 재설정 실패: " + str(error))
    st.stop()

try:
    if preview:
        row = {"body": json.loads((HERE / "seed.json").read_text(encoding="utf-8")), "revision": 0}
        actor = {"display_name": "Hugo", "color": "#4285f4"}
        members = [{"display_name": "Hugo", "color": "#4285f4"}, {"display_name": "Jayna", "color": "#b278ed"}]
        st.info("화면 미리보기 · 데이터 변경은 저장되지 않습니다.")
        token = None
    else:
        token = session_token()
        if st.session_state.get("remember_login"):
            refresh = st.session_state.auth["refresh_token"]
            if st.session_state.login_storage_command.get("token") != refresh:
                queue_login_storage("store", refresh)
                st.rerun()
        user = api("/auth/v1/user", token=token)
        members = api("/rest/v1/operation_members?select=email,display_name,color&order=display_name", token=token)
        actor = next((m for m in members if m["email"].lower() == user["email"].lower()), None)
        if not actor:
            st.error("이 업무판에 등록된 팀원 계정이 아닙니다.")
            if st.button("다른 계정으로 로그인"):
                st.session_state.clear()
                st.rerun()
            st.stop()
        with st.sidebar:
            st.write(f"● {actor['display_name']}")
            st.caption("화면은 5초마다 최신 내용을 확인합니다.")
            if st.button("로그아웃"):
                try:
                    api("/auth/v1/logout", "POST", token=token)
                finally:
                    st.session_state.clear()
                    st.session_state.login_storage_ready = True
                    queue_login_storage("clear")
                    st.rerun()
            if st.button("새로고침"):
                st.rerun()
        rows = api("/rest/v1/operation_boards?id=eq.operation&select=body,revision", token=token)
        if not rows:
            st.error("업무판 초기 데이터가 아직 연결되지 않았습니다.")
            st.stop()
        row = rows[0]
except (RuntimeError, urllib.error.URLError):
    st.error("공동 저장 서버에 연결하지 못했어요. 변경하지 않고 연결을 다시 확인해주세요.")
    if st.button("다시 연결"):
        st.rerun()
    st.stop()

if not preview:
    with st.sidebar.expander("진행 중 오더 · 원본 업데이트"):
        order_file = st.file_uploader("Open SO Status 엑셀", type=["xlsx", "xlsm"], key="order_file")
        source_link = st.text_input("원본 엑셀 링크", value=row["body"].get("orderSummary", {}).get("sourceUrl", ""))
        if st.button("오더 요약 업데이트", disabled=order_file is None):
            try:
                source_link = source_link.strip()
                if source_link and not source_link.startswith("https://"):
                    raise ValueError("원본 링크는 https 주소로 입력해주세요.")
                summary = parse_order_workbook(order_file.getvalue(), order_file.name, source_link)
                for attempt in range(3):
                    latest = api("/rest/v1/operation_boards?id=eq.operation&select=body,revision", token=token)[0]
                    latest["body"]["orderSummary"] = summary
                    saved = api("/rest/v1/rpc/operation_save", "POST", {
                        "p_expected_revision": latest["revision"], "p_body": latest["body"],
                        "p_request_id": str(uuid.uuid4()), "p_message": "진행 중 오더 요약을 업데이트했어요",
                    }, token=token)
                    if saved.get("ok"):
                        break
                if not saved.get("ok"):
                    raise ValueError("다른 변경과 겹쳐 저장하지 못했어요. 다시 눌러주세요.")
                st.session_state.order_import_success = len(summary["orders"])
                st.rerun()
            except (ValueError, RuntimeError, KeyError, OSError, SyntaxError) as error:
                st.error("오더 업데이트 실패: " + str(error))
        if "order_import_success" in st.session_state:
            st.success(f"진행 중 {st.session_state.order_import_success}건 저장됨")

result = board_component(BOARD_HTML, BOARD_JS, BOARD_CSS)(
    data={"board": row["body"], "revision": row["revision"], "actor": actor,
          "members": members, "ack": st.session_state.get("ack"), "readOnly": preview},
    key="operation", on_change_change=lambda: None, on_refresh_change=lambda: None,
)
change = result.get("change")
if change and change.get("id") != st.session_state.get("handled_request"):
    request_id = change["id"]
    try:
        if preview:
            response = {"ok": False, "message": "미리보기에서는 저장되지 않습니다."}
        else:
            response = api("/rest/v1/rpc/operation_save", "POST", {
                "p_expected_revision": change["revision"], "p_body": change["board"],
                "p_request_id": request_id, "p_message": change.get("message", "업무판을 업데이트했어요"),
            }, token=token)
        st.session_state.ack = {"id": request_id, **response}
    except (RuntimeError, urllib.error.URLError):
        st.session_state.ack = {"id": request_id, "ok": False, "message": "저장 실패. 다시 시도해주세요."}
    st.session_state.handled_request = request_id
    st.rerun()
