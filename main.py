# ============================================================
# xing_watch_and_theme_ocr.py
# ✅ 한 파일 (메인은 Python 3.14-32/64 whichever has XING COM 등록된 쪽에서 실행)
# ✅ (1) XING 로그인 → 조건검색 '관심종목' 전체 로드 (3.14)
# ✅ (2) 이미지 선택 → THEME OCR(너 코드 그대로) 는 Python 3.12 subprocess로 실행 (cv2/easyocr 설치된 쪽)
# ✅ (3) CMD에 테마 출력 + 관심종목 코드풀 출력 (매핑은 아직 안 함)
#
# 실행 예)
#   py -3.14 C:\Users\User\Desktop\xing_watch_and_theme_ocr.py
#
# 필요)
# - 3.14 환경: XING COM(XA_Session, XA_DataSet) "클래스 등록" 되어 있어야 함
# - 3.12 환경: pip install opencv-python easyocr python-Levenshtein numpy
# ============================================================

import os
import time
import re
import json
import tempfile
import subprocess
import pythoncom
import win32com.client
from dataclasses import dataclass

# =========================
# OCR용 Python(3.12) 지정
# =========================
PY312_EXE = r"C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe"
USE_PY_LAUNCHER = False  # True면 ["py","-3.12"]

# =========================
# XING 설정
# =========================
@dataclass
class XingConfig:
    user_id: str = os.environ.get("XING_USER_ID", "")
    user_pw: str = os.environ.get("XING_USER_PW", "")
    cert_pw: str = os.environ.get("XING_CERT_PW", "")
    server: str = os.environ.get("XING_SERVER", "real")  # real/demo
    timeout_sec: int = 12

CFG = XingConfig()

RES_DIR = r"C:\xingAPI_Program(2025.06.07)\Res"
SERVER_ADDR = {"real": "hts.ebestsec.co.kr", "demo": "demo.ebestsec.co.kr"}

COND_WATCH = "관심종목"
COOLDOWN_SEC = 1.2
RETRY_MAX = 1
RETRY_SLEEP_SEC = 1.5

# =========================
# XING 유틸
# =========================
def sstrip(x) -> str:
    return (x or "").strip()

def to_float_or_none(x):
    s = sstrip(x).replace(",", "")
    if s == "" or s == "-":
        return None
    try:
        return float(s)
    except Exception:
        return None

def is_etn(code: str) -> bool:
    return (not code) or (not code.isdigit())

def is_etf(name: str) -> bool:
    if not name:
        return True
    u = name.upper()
    return any(k in u for k in ["KODEX", "TIGER", "KBSTAR", "ARIRANG", "HANARO", "ACE", "SOL", "ETF", "ETN"])

def is_etf_etn(code: str, name: str) -> bool:
    return is_etn(code) or is_etf(name)

def sort_by_rate_desc(rows):
    def key(r):
        v = r.get("rate")
        return (-1e18 if v is None else v)
    return sorted(rows, key=key, reverse=True)

# =========================
# XING 이벤트
# =========================
class XASessionEvents:
    def OnLogin(self, code, msg):
        self.parent._login_code = code
        self.parent._login_msg = msg

class XAQueryEvents:
    def OnReceiveData(self, tr_code):
        self.parent._received = True
        self.parent._last_tr = tr_code

# =========================
# XING API
# =========================
class XingAPI:
    def __init__(self):
        self._received = False
        self._last_tr = ""
        self._login_code = None
        self._login_msg = ""

        self.session = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEvents)
        self.session.parent = self

        self.query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
        self.query.parent = self

    def _wait(self, timeout, tag="TR timeout"):
        st = time.time()
        while not self._received:
            pythoncom.PumpWaitingMessages()
            if time.time() - st > timeout:
                raise TimeoutError(tag)
            time.sleep(0.01)

    def _set_res(self, res_filename: str) -> str:
        path = os.path.join(RES_DIR, res_filename)
        if not os.path.exists(path):
            raise FileNotFoundError(f"res 파일 없음: {path}")
        self.query.ResFileName = path
        return path

    def _request_service_or_request0(self, tr_code: str):
        if hasattr(self.query, "RequestService"):
            return self.query.RequestService(tr_code, "")
        return self.query.Request(0)

    def login(self):
        if not CFG.user_id or not CFG.user_pw or not CFG.cert_pw:
            raise RuntimeError("❌ 환경변수 XING_USER_ID / XING_USER_PW / XING_CERT_PW 설정 필요")

        addr = SERVER_ADDR[CFG.server]
        if not self.session.ConnectServer(addr, 20001):
            raise RuntimeError("서버 연결 실패")

        server_type = 0 if CFG.server == "real" else 1
        self.session.Login(CFG.user_id, CFG.user_pw, CFG.cert_pw, server_type, 0)

        st = time.time()
        while self._login_code is None:
            pythoncom.PumpWaitingMessages()
            if time.time() - st > CFG.timeout_sec:
                raise TimeoutError("로그인 응답 타임아웃")
            time.sleep(0.01)

        if self._login_code != "0000":
            raise RuntimeError(f"로그인 실패: {self._login_code} {self._login_msg}")

        print("[LOGIN] 성공")

    def t1866_list(self):
        self._received = False
        self._set_res("t1866.res")

        inb = "t1866InBlock"
        self.query.SetFieldData(inb, "user_id", 0, CFG.user_id)
        self.query.SetFieldData(inb, "gb", 0, "0")
        self.query.SetFieldData(inb, "group_name", 0, "")
        self.query.SetFieldData(inb, "cont", 0, "0")
        self.query.SetFieldData(inb, "cont_key", 0, "")

        ret = self.query.Request(0)
        if ret < 0:
            raise RuntimeError(f"t1866 Request 실패 ret={ret}")
        self._wait(CFG.timeout_sec, "t1866 timeout")

        outb = "t1866OutBlock1"
        cnt = self.query.GetBlockCount(outb)

        rows = []
        for i in range(cnt):
            rows.append({
                "query_index": sstrip(self.query.GetFieldData(outb, "query_index", i)),
                "query_name": sstrip(self.query.GetFieldData(outb, "query_name", i)),
            })
        return rows

    def t1857_snapshot_S0(self, query_index: str):
        self._set_res("t1857.res")

        inb = "t1857InBlock"
        self.query.SetFieldData(inb, "sRealFlag", 0, "0")
        self.query.SetFieldData(inb, "sSearchFlag", 0, "S")
        self.query.SetFieldData(inb, "query_index", 0, str(query_index))

        last_ret = None
        for _ in range(RETRY_MAX):
            time.sleep(COOLDOWN_SEC)
            self._received = False
            ret = self._request_service_or_request0("t1857")
            last_ret = ret
            if ret >= 0:
                self._wait(CFG.timeout_sec, "t1857 timeout")
                break
            time.sleep(RETRY_SLEEP_SEC)

        if last_ret is None or last_ret < 0:
            raise RuntimeError(f"t1857 호출 실패 ret={last_ret}")

        outb = "t1857OutBlock1"
        cnt = self.query.GetBlockCount(outb)

        rows = []
        for i in range(cnt):
            code = sstrip(self.query.GetFieldData(outb, "shcode", i))
            name = sstrip(self.query.GetFieldData(outb, "hname", i))
            rate = to_float_or_none(self.query.GetFieldData(outb, "diff", i))
            if code and not is_etf_etn(code, name):
                rows.append({"code": code, "name": name, "rate": rate})
        return sort_by_rate_desc(rows)

# =========================
# OCR subprocess (3.12): 테마 추출만
# =========================
THEME_POOL = [
    "로봇","휴머노이드","의료AI","자율주행","감속기",
    "이재명","정치","서울시장","대선","정원오","지역화폐",
    "전기차","부동산","자산","전력",
    "항암","면역항암제","면역항암","RNA","mRNA",
    "비만","비만치료제","위고비","일라이릴리",
    "엔비디아","테슬라","스페이스x","스페이스",
    "메타버스","LCD","OLED","원격","의료기기","비대면",
    "반도체","VR","가상현실","가상","파운드리",
    "유전자","진단키트","키트",
    "AI","서버","ESS","온디바이스","양자","양자컴퓨터",
    "폴더블","폴더블폰","터치패널","강화유리",
    "휴대폰","스마트폰","칩","통신",
    "게임","엔터","빌보드",
    "안테나","보안","해킹",
    "삼성","SK하이닉스","SK","지주사","반도체","2차전지",
    "증권","스테이블","스테이블코인","IOT","건설","철도",
    "당뇨","치매","알츠하이머","치매약",
    "전해질","음극재","양극재","음극","양극",
    "방산","국방","미사일",
    "재건","우크라이나","우크라이나재건","전쟁",
    "우주","우주항공","조선","LNG","LPG","엔진",
    "바이오","제약","줄기세포","마이크로바이옴",
    "리튬","탄소포집","탄소","자동차",
    "5G","6G","통신장비",
    "신규주","중동재건","남북경협",
    "헷지","헷지주","총선","저출산",
    "결제","가상화폐","비트코인",
    "보험","은행","금융",
    "카지노","호텔","여행",
    "뷰티","K뷰티","화장품",
    "음식","식자재","프랜차이즈",
    "웹툰","미디어","광고",
    "이커머스","유통","콘솔",
    "그래핀","치아","화학","세라믹","기계",
    "희토류","구리","알루미늄","철강",
    "금","은","금속",
    "자원","해저자원","위성","로켓","드론",
    "태양광","풍력","태양광발전","풍력발전",
    "원전","재활용",
    "기후","온난화","지구온난화"
]

def _get_py312_cmd():
    if USE_PY_LAUNCHER:
        return ["py", "-3.12"]
    if not os.path.exists(PY312_EXE):
        raise RuntimeError(f"PY312 경로가 잘못됨: {PY312_EXE}")
    return [PY312_EXE]

def run_theme_ocr_py312_return_themes():
    """
    ✅ 3.12에서 cv2/easyocr로 테마만 추출
    ✅ stdout은 bytes로 받고 UTF-8/CP949 자동 복구
    ✅ 마지막에 JSON 한 덩어리만 출력하도록 서브스크립트 설계
    """
    pycmd = _get_py312_cmd()

    ocr_script = r"""
import json, re
from tkinter import Tk, filedialog

import cv2
import numpy as np
import easyocr
import Levenshtein

THEME_POOL = __THEME_POOL__

def h2j(text):
    CHO = ['ㄱ','ㄲ','ㄴ','ㄷ','ㄸ','ㄹ','ㅁ','ㅂ','ㅃ','ㅅ','ㅆ','ㅇ','ㅈ','ㅉ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ']
    JUNG = ['ㅏ','ㅐ','ㅑ','ㅒ','ㅓ','ㅔ','ㅕ','ㅖ','ㅗ','ㅘ','ㅙ','ㅚ','ㅛ','ㅜ','ㅝ','ㅞ','ㅟ','ㅠ','ㅡ','ㅢ','ㅣ']
    JONG = ['','ㄱ','ㄲ','ㄳ','ㄴ','ㄵ','ㄶ','ㄷ','ㄹ','ㄺ','ㄻ','ㄼ','ㄽ','ㄾ','ㄿ','ㅀ','ㅁ','ㅂ','ㅄ','ㅅ','ㅆ','ㅇ','ㅈ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ']
    res = ""
    for c in text:
        if '가' <= c <= '힣':
            code = ord(c) - ord('가')
            res += CHO[code//588] + JUNG[(code//28)%21] + JONG[code%28]
        else:
            res += c
    return res

def correct_theme_from_pool(raw):
    clean = re.sub(r'[^가-힣A-Z0-9]', '', (raw or "").upper())
    if len(clean) < 2:
        return None

    # exact
    for t in THEME_POOL:
        if t.upper() == clean:
            return t

    cj = h2j(clean)
    best, best_sim = None, 0.0
    for t in THEME_POOL:
        tj = h2j(t.upper())
        dist = Levenshtein.distance(cj, tj)
        sim = 1 - dist / max(len(cj), len(tj))
        if t.startswith(clean[:1]):
            sim += 0.2
        if sim > best_sim:
            best_sim = sim
            best = t
    return best if best_sim >= 0.5 else None

def pick_images():
    root = Tk()
    root.withdraw()
    files = filedialog.askopenfilenames(
        title="분류표(테마) 캡처 이미지 선택",
        filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.webp"), ("All files", "*.*")]
    )
    root.destroy()
    return list(files)

def extract_themes_from_one(img_path, reader):
    img = cv2.imread(img_path)
    if img is None:
        return []

    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    mask = cv2.inRange(hsv, (15,70,120), (45,255,255))
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    out = []
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if w < 30 or h < 8:
            continue
        roi = img[y:y+h, x:x+w]
        roi = cv2.resize(roi, None, fx=3, fy=3)
        raw = "".join(reader.readtext(roi, detail=0))
        corrected = correct_theme_from_pool(raw)
        if corrected:
            out.append(corrected)
    return out

def main():
    files = pick_images()
    if not files:
        print(json.dumps({"ok": True, "themes": [], "files": []}, ensure_ascii=False))
        return 0

    # GPU는 환경마다 불안정할 수 있어 기본 CPU
    reader = easyocr.Reader(['ko','en'], gpu=False)

    found = set()
    for p in files:
        for t in extract_themes_from_one(p, reader):
            found.add(t)

    themes = sorted(found)
    print(json.dumps({"ok": True, "themes": themes, "files": files}, ensure_ascii=False))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
"""

    # 리스트 주입(안전하게 json으로)
    injected = ocr_script.replace("__THEME_POOL__", json.dumps(THEME_POOL, ensure_ascii=False))

    with tempfile.NamedTemporaryFile("w", suffix=".py", delete=False, encoding="utf-8") as tf:
        tf.write(injected)
        tmp_path = tf.name

    try:
        r = subprocess.run(pycmd + [tmp_path], capture_output=True)
        raw = r.stdout or b""
        if not raw:
            # stderr에도 있을 수 있음
            raw = (r.stderr or b"")

        # ✅ 인코딩 자동 복구
        text = None
        for enc in ("utf-8", "cp949", "euc-kr"):
            try:
                text = raw.decode(enc, errors="strict")
                break
            except Exception:
                continue
        if text is None:
            text = raw.decode("utf-8", errors="ignore")

        # ✅ JSON blob만 추출
        blob = text.strip()
        if not (blob.startswith("{") and blob.endswith("}")):
            s = text.find("{")
            e = text.rfind("}")
            if s != -1 and e != -1 and e > s:
                blob = text[s:e+1]

        j = json.loads(blob)
        if not isinstance(j, dict) or not j.get("ok"):
            return []

        themes = j.get("themes", [])
        # 보험: 한글 포함 테마만
        cleaned = []
        for t in themes:
            if isinstance(t, str) and re.search(r"[가-힣]", t):
                cleaned.append(t)
        return cleaned

    finally:
        try:
            os.remove(tmp_path)
        except Exception:
            pass

# =========================
# main
# =========================
def main():
    pythoncom.CoInitialize()

    # 1) XING 관심종목 로드 (3.14에서만 성공해야 함)
    watch_rows = []
    try:
        x = XingAPI()
        x.login()

        conds = x.t1866_list()
        name_to_qidx = {c["query_name"]: c["query_index"] for c in conds if c.get("query_name") and c.get("query_index")}
        watch_qidx = name_to_qidx.get(COND_WATCH, "")

        if not watch_qidx:
            print("❌ 조건검색 목록에서 '관심종목'을 못 찾음")
            print("현재 조건검색 이름들:", [c.get("query_name") for c in conds])
        else:
            watch_rows = x.t1857_snapshot_S0(watch_qidx)
            print(f"\n[OK] 관심종목 로드 {len(watch_rows)}개")

    except Exception as e:
        # XING 실패해도 OCR은 진행 가능
        print(f"\n[XING] 실패 -> 테마 OCR만 진행: {e}")

    # 2) 테마 OCR (3.12 subprocess)
    print("\n[STEP] 분류표(테마) 이미지 선택 + OCR(3.12)...")
    themes = []
    try:
        themes = run_theme_ocr_py312_return_themes()
    except Exception as e:
        print(f"❌ 테마 OCR 실패: {e}")
        themes = []

    # 3) 출력
    print("\n===============================")
    print("✅ 최종 추출 테마 목록")
    print("===============================")
    if themes:
        for t in themes:
            print(f"- {t}")
    else:
        print("(없음)")
    print("===============================")
    print("총 테마 개수:", len(themes))

    if watch_rows:
        print("\n===============================")
        print("✅ 관심종목 코드풀 (상위 200개만)")
        print("===============================")
        for r in watch_rows[:200]:
            rt = "?" if r["rate"] is None else f'{r["rate"]:.2f}%'
            print(f'{r["code"]}  {r["name"]}  {rt}')
        print("===============================")
        print("총 관심종목 개수:", len(watch_rows))
    else:
        print("\n[WARN] 관심종목을 못 가져와서 코드풀 출력은 생략")

if __name__ == "__main__":
    main()
