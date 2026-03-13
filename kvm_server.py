"""
==================================================
  Software KVM - 메인 PC (송신 측) v5.0
==================================================
설치: pip install pyautogui pyperclip screeninfo
설정: kvm_config.ini 를 수정하세요
실행: python kvm_server.py  (UAC 자동 승격)
==================================================

[설계]
  pynput 완전 제거 - pynput과 Windows LL훅 충돌 문제 해결
  Windows SetWindowsHookEx(WH_KEYBOARD_LL / WH_MOUSE_LL) 로만 동작
  메인 스레드에서 GetMessage 루프 실행 (훅은 메인 스레드 필수)
==================================================
"""

import socket, json, threading, time, sys, os, configparser, ctypes, ctypes.wintypes
import pyautogui, pyperclip

pyautogui.FAILSAFE = False
pyautogui.PAUSE    = 0

# ──────────────────────────────────────────────
# UAC 자동 승격
# ──────────────────────────────────────────────
def ensure_admin():
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        is_admin = False
    if not is_admin:
        print("[UAC] 관리자 권한 필요 - UAC 승격 요청 중...")
        ret = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable,
            " ".join(f'"{a}"' for a in sys.argv),
            None, 1)
        if ret > 32:
            sys.exit(0)
        else:
            print(f"[ERR] UAC 승격 실패 (ret={ret}). 관리자 권한으로 직접 실행하세요.")
            input("Enter 키를 눌러 종료...")
            sys.exit(1)

ensure_admin()

# ──────────────────────────────────────────────
# 설정 파일 로드
# ──────────────────────────────────────────────
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kvm_config.ini")

def load_config():
    cfg = configparser.ConfigParser()
    if not os.path.exists(CONFIG_FILE):
        print(f"[ERR] 설정 파일 없음: {CONFIG_FILE}")
        input("Enter 키를 눌러 종료...")
        sys.exit(1)
    cfg.read(CONFIG_FILE, encoding='utf-8')
    return cfg

cfg = load_config()
SECOND_PC_IP       = cfg.get        ('network',  'second_pc_ip',      fallback='192.168.0.XXX')
PORT               = cfg.getint     ('network',  'port',               fallback=9999)
RECONNECT_SEC      = cfg.getint     ('network',  'reconnect_sec',      fallback=3)
SECOND_PC_POSITION = cfg.get        ('layout',   'second_pc_position', fallback='left')
EDGE_THRESHOLD     = cfg.getint     ('layout',   'edge_threshold',     fallback=8)
CLIPBOARD_SYNC     = cfg.getboolean ('features', 'clipboard_sync',     fallback=True)
VERBOSE            = cfg.getboolean ('log',      'verbose',            fallback=True)
LOG_MOUSE_MOVE     = cfg.getboolean ('log',      'log_mouse_move',     fallback=False)
LOG_PACKET_RATE    = cfg.getboolean ('log',      'log_packet_rate',    fallback=False)
MOUSE_SKIP_PACKETS = cfg.getboolean ('features', 'mouse_skip_packets', fallback=True)

# ──────────────────────────────────────────────
# 로깅
# ──────────────────────────────────────────────
_log_lock = threading.Lock()

def _ts():
    t = time.time()
    return time.strftime('%H:%M:%S.') + f'{int(t*1000)%1000:03d}'

def log(msg):
    if VERBOSE:
        with _log_lock:
            print(f"[{_ts()}] {msg}", flush=True)

def log_info(msg):
    with _log_lock:
        print(f"[{_ts()}] {msg}", flush=True)

# ──────────────────────────────────────────────
# 다중 모니터 가상 화면
# ──────────────────────────────────────────────
def get_virtual_screen():
    try:
        import screeninfo
        monitors = screeninfo.get_monitors()
        L = min(m.x for m in monitors)
        T = min(m.y for m in monitors)
        R = max(m.x + m.width  for m in monitors)
        B = max(m.y + m.height for m in monitors)
        return L, T, R, B, monitors
    except Exception as e:
        log_info(f"[WARN] screeninfo 실패({e}), 단일 모니터로 처리")
        w, h = pyautogui.size()
        return 0, 0, w, h, None

VS_LEFT, VS_TOP, VS_RIGHT, VS_BOTTOM, MONITORS = get_virtual_screen()
VS_W = max(VS_RIGHT  - VS_LEFT, 1)
VS_H = max(VS_BOTTOM - VS_TOP,  1)

# ──────────────────────────────────────────────
# Windows API 설정
# ──────────────────────────────────────────────
user32   = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

WH_KEYBOARD_LL = 13
WH_MOUSE_LL    = 14
HC_ACTION      = 0
WM_QUIT        = 0x0012
WM_MOUSEMOVE   = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP   = 0x0202
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP   = 0x0205
WM_MBUTTONDOWN = 0x0207
WM_MBUTTONUP   = 0x0208
WM_MOUSEWHEEL  = 0x020A
WM_KEYDOWN     = 0x0100
WM_KEYUP       = 0x0101
WM_SYSKEYDOWN  = 0x0104
WM_SYSKEYUP    = 0x0105
LLMHF_INJECTED = 0x00000001

ULONG_PTR = ctypes.c_uint64 if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_uint32

HOOKPROC = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_int, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM)

user32.SetWindowsHookExW.restype  = ctypes.c_void_p
user32.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC, ctypes.c_void_p, ctypes.wintypes.DWORD]
user32.CallNextHookEx.restype     = ctypes.c_long
user32.CallNextHookEx.argtypes    = [ctypes.c_void_p, ctypes.c_int, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM]
user32.UnhookWindowsHookEx.restype  = ctypes.wintypes.BOOL
user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]

class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [("vkCode", ctypes.wintypes.DWORD), ("scanCode", ctypes.wintypes.DWORD),
                ("flags", ctypes.wintypes.DWORD),  ("time", ctypes.wintypes.DWORD),
                ("dwExtraInfo", ULONG_PTR)]

class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

class MSLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [("pt", POINT), ("mouseData", ctypes.wintypes.DWORD),
                ("flags", ctypes.wintypes.DWORD), ("time", ctypes.wintypes.DWORD),
                ("dwExtraInfo", ULONG_PTR)]

# ──────────────────────────────────────────────
# 전역 상태
# ──────────────────────────────────────────────
import queue as _queue

sock              = None
sock_lock         = threading.Lock()
active_pc         = "main"
switching         = False
_block_input      = False
_last_clip        = ""
_second_screen_w  = 1920
_second_screen_h  = 1080
_kb_hook_id       = None
_mouse_hook_id    = None
_hook_thread_id   = None   # 훅이 설치된 스레드 ID

# 송신 큐: 훅 콜백에서 직접 소켓 I/O 금지 → 큐에 넣고 별도 스레드가 전송
_send_queue       = _queue.Queue(maxsize=512)

def _send_worker():
    """큐에서 이벤트를 꺼내 소켓으로 전송하는 전용 스레드"""
    global sock
    while True:
        try:
            event = _send_queue.get(timeout=1.0)
            if event is None:
                break
            with sock_lock:
                s = sock
            if s is None:
                continue
            try:
                s.sendall((json.dumps(event) + '\n').encode('utf-8'))
            except OSError as e:
                with sock_lock:
                    sock = None
                log_info(f"[WARN] 전송 실패: {e}")
        except _queue.Empty:
            continue
        except Exception as e:
            log(f"[ERR] send_worker: {e}")

# 윈도우키 VK 코드 집합 (특별 차단 처리 필요)
_WIN_KEYS = {0x5B, 0x5C}

# extended key가 필요한 vk 코드 (KEYEVENTF_EXTENDEDKEY 플래그 필요)
_EXTENDED_KEYS = {
    0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27, 0x28,  # PageUp/Dn, End, Home, 방향키
    0x2D, 0x2E,        # Insert, Delete
    0x5B, 0x5C,        # Win키
    0xA3,              # Ctrl_r
    0xA5,              # Alt_r
}

# ──────────────────────────────────────────────
# 엣지 감지 조건
# ──────────────────────────────────────────────
# 엣지 감지 조건 및 전환 위치 계산
# second_pos: 세컨드 PC가 메인의 어느 방향에 있는지
#
# to_second_cond(x,y): 메인에서 세컨드로 전환할 엣지 조건
# to_main_cond(x,y)  : 세컨드에서 메인으로 복귀할 엣지 조건
# entry_for_second(main_x, main_y, s_w, s_h):
#   세컨드 진입 시 세컨드 커서 시작 위치 (논리 픽셀)
#   → 연속된 화면처럼 세컨드 화면의 반대편 끝에서 시작
# park_for_second(main_x, main_y):
#   세컨드 제어 중 메인 마우스를 어디에 주차할지
#   → 복귀 트리거(반대편 끝) 바로 안쪽에 주차
# return_for_main(second_x, second_y, s_w, s_h):
#   세컨드→메인 복귀 시 메인 커서 위치
#   → 진입했던 쪽 끝, y는 세컨드 y 비율로

T = EDGE_THRESHOLD
_pos = SECOND_PC_POSITION

# to_main_cond_second: 세컨드 커서(논리좌표) 기준 복귀 조건
# 전역 _second_screen_w/h는 DPI 수신 후 업데이트되므로 함수에서 직접 참조
def to_main_cond_second(sx, sy):
    if _pos == "left":   return sx >= _second_screen_w - 3
    if _pos == "right":  return sx <= 2
    if _pos == "top":    return sy >= _second_screen_h - 3
    if _pos == "bottom": return sy <= 2
    return False

# 주차 위치: 항상 메인 화면 중앙
# 이유: 어느 방향에서 진입해도 delta 계산이 안정적
#       화면 끝에 주차하면 한 방향 이동이 막혀 세컨드 커서가 달라붙음
_PARK_CENTER = lambda mx,my: ((VS_LEFT + VS_RIGHT) // 2, (VS_TOP + VS_BOTTOM) // 2)

if _pos == "left":
    to_second_cond   = lambda x,y: x <= VS_LEFT  + T
    to_main_cond     = lambda x,y: x >= VS_RIGHT - T
    entry_for_second = lambda mx,my,sw,sh: (sw - 2, int((my - VS_TOP) / VS_H * sh))
    park_for_second  = _PARK_CENTER
    return_for_main  = lambda sx,sy,sw,sh: (VS_LEFT + T*2, VS_TOP + int(sy / sh * VS_H))

elif _pos == "right":
    to_second_cond   = lambda x,y: x >= VS_RIGHT - T
    to_main_cond     = lambda x,y: x <= VS_LEFT  + T
    entry_for_second = lambda mx,my,sw,sh: (1, int((my - VS_TOP) / VS_H * sh))
    park_for_second  = _PARK_CENTER
    return_for_main  = lambda sx,sy,sw,sh: (VS_RIGHT - T*2, VS_TOP + int(sy / sh * VS_H))

elif _pos == "top":
    to_second_cond   = lambda x,y: y <= VS_TOP    + T
    to_main_cond     = lambda x,y: y >= VS_BOTTOM - T
    entry_for_second = lambda mx,my,sw,sh: (int((mx - VS_LEFT) / VS_W * sw), sh - 2)
    park_for_second  = _PARK_CENTER
    return_for_main  = lambda sx,sy,sw,sh: (VS_LEFT + int(sx / sw * VS_W), VS_TOP + T*2)

else:  # bottom
    to_second_cond   = lambda x,y: y >= VS_BOTTOM - T
    to_main_cond     = lambda x,y: y <= VS_TOP    + T
    entry_for_second = lambda mx,my,sw,sh: (int((mx - VS_LEFT) / VS_W * sw), 1)
    park_for_second  = _PARK_CENTER
    return_for_main  = lambda sx,sy,sw,sh: (VS_LEFT + int(sx / sw * VS_W), VS_BOTTOM - T*2)

# ──────────────────────────────────────────────
# 네트워크
# ──────────────────────────────────────────────
def _packet_rate_logger():
    """1초마다 송신 패킷 수를 로그로 출력 (LOG_PACKET_RATE=true일 때만)"""
    global _pkt_sent_count
    while True:
        time.sleep(1.0)
        if LOG_PACKET_RATE and active_pc == "second":
            with _pkt_sent_lock:
                count = _pkt_sent_count
                _pkt_sent_count = 0
            log_info(f"[PKT-RATE] 송신: {count} pkt/s  큐잔량: {_send_queue.qsize()}")
        else:
            with _pkt_sent_lock:
                _pkt_sent_count = 0

def send_event(event: dict):
    """훅 콜백에서 호출 가능 - 소켓 I/O 없이 큐에만 넣음 (논블로킹)"""
    if sock is None:
        return
    try:
        _send_queue.put_nowait(event)
    except _queue.Full:
        pass  # 큐가 꽉 찼으면 드롭 (마우스 이동 과잉 패킷 방지)

def connect_loop():
    global sock
    while True:
        if sock is None:
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(5)
                s.connect((SECOND_PC_IP, PORT))
                s.settimeout(None)
                s.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1)
                with sock_lock:
                    sock = s
                log_info(f"[OK] 세컨드 PC 연결: {SECOND_PC_IP}:{PORT}")
                if CLIPBOARD_SYNC:
                    _push_clipboard()
                send_event({'type': 'dpi_request'})
            except OSError:
                log_info(f"[..] 재연결 대기 ({SECOND_PC_IP}:{PORT})")
        time.sleep(RECONNECT_SEC)

# 비상 탈출 키 추적: Ctrl+Alt+F12 동시 감지
_emergency_keys_down = set()
EMERGENCY_COMBO = {0x11, 0x12, 0x7B}  # VK_CTRL=0x11, VK_ALT=0x12, VK_F12=0x7B

def _emergency_release():
    """비상 탈출: 모든 차단 해제 후 메인 PC 복귀"""
    global active_pc, switching, _block_input
    log_info("[EMERGENCY] ★★★ 비상 탈출! Ctrl+Alt+F12 ★★★")
    # 즉시 모든 상태 강제 초기화
    switching    = False
    _block_input = False
    active_pc    = "main"
    # 커서 복원
    try:
        _restore_cursor()
    except Exception as e:
        log_info(f"[EMERGENCY] 커서 복원 실패: {e}")
    log_info("[EMERGENCY] 완료 - 메인 PC 제어 복귀")

def recv_loop():
    global sock, _last_clip, _second_screen_w, _second_screen_h, _second_phys_w, _second_phys_h
    buf = ""
    while True:
        with sock_lock:
            s = sock
        if s is None:
            time.sleep(0.5); buf = ""; continue
        try:
            data = s.recv(4096).decode('utf-8', errors='ignore')
            if not data:
                time.sleep(0.05); continue
            buf += data
            while '\n' in buf:
                line, buf = buf.split('\n', 1)
                line = line.strip()
                if not line: continue
                try:
                    ev = json.loads(line)
                    t  = ev.get('type')
                    if t == 'clipboard':
                        text = ev.get('text', '')
                        if text and text != _last_clip:
                            _last_clip = text
                            pyperclip.copy(text)
                            log_info(f"[CB] <- 세컨드: {text[:60]}")
                    elif t == 'dpi_info':
                        phys_w = ev.get('width',  1920)
                        phys_h = ev.get('height', 1080)
                        scale  = ev.get('scale',  1.0)
                        # logical_w/h: 클라이언트가 pyautogui.moveTo()에 쓸 논리 해상도
                        # 없으면 물리/scale 로 계산
                        _second_screen_w = ev.get('logical_w', int(phys_w / scale) if scale > 0 else phys_w)
                        _second_screen_h = ev.get('logical_h', int(phys_h / scale) if scale > 0 else phys_h)
                        _second_phys_w   = phys_w
                        _second_phys_h   = phys_h
                        log_info(f"[DPI] 세컨드 화면: 물리={phys_w}x{phys_h}, scale={scale:.2f}x, 논리(전송좌표)={_second_screen_w}x{_second_screen_h}")
                except json.JSONDecodeError:
                    pass
        except OSError:
            # 세컨드 PC 연결 끊어짐 → 자동으로 메인 PC 복귀
            if active_pc == "second":
                log_info("[NET] 세컨드 PC 연결 끊어짐 - 자동 메인 복귀")
                threading.Thread(target=_emergency_release, daemon=True).start()
            time.sleep(0.5); buf = ""

# ──────────────────────────────────────────────
# 클립보드
# ──────────────────────────────────────────────
def _push_clipboard():
    global _last_clip
    try:
        text = pyperclip.paste()
        if text and text != _last_clip:
            _last_clip = text
            send_event({'type': 'clipboard', 'text': text})
            log_info(f"[CB] -> 세컨드: {text[:60]}")
    except Exception:
        pass

def _pull_clipboard():
    send_event({'type': 'clipboard_request'})

# ──────────────────────────────────────────────
# 투명 커서 생성 (커서 숨김용)
# ──────────────────────────────────────────────
# CreateCursor: 1x1 투명 커서 생성
# ShowCursor는 스레드별 카운터라 훅 스레드 밖에서 호출하면 효과 없음
# SetSystemCursor로 시스템 커서를 투명 커서로 교체하는 방식 사용

_original_cursors = {}   # 원본 커서 핸들 백업
_OCR_NORMAL = 32512      # IDC_ARROW

def _make_transparent_cursor():
    """1x1 투명 커서 생성"""
    try:
        # AND 마스크=0xFF(투명), XOR 마스크=0x00(검정)
        and_mask = (ctypes.c_ubyte * 4)(0xFF, 0xFF, 0xFF, 0xFF)
        xor_mask = (ctypes.c_ubyte * 4)(0x00, 0x00, 0x00, 0x00)
        hcursor = user32.CreateCursor(
            None, 0, 0, 1, 1,
            ctypes.cast(and_mask, ctypes.c_void_p),
            ctypes.cast(xor_mask, ctypes.c_void_p)
        )
        return hcursor
    except Exception as e:
        log_info(f"[WARN] 투명 커서 생성 실패: {e}")
        return None

# OCR(커서 종류) 목록 - 주요 커서들을 모두 교체
_CURSOR_IDS = [32512, 32513, 32514, 32515, 32516, 32640, 32641, 32642,
               32643, 32644, 32645, 32646, 32648, 32649, 32650, 32651]

# CopyCursor는 Win32 매크로(=CopyIcon) - user32에 직접 없음
user32.CopyIcon.restype    = ctypes.c_void_p
user32.CopyIcon.argtypes   = [ctypes.c_void_p]
user32.LoadCursorW.restype  = ctypes.c_void_p
user32.LoadCursorW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
user32.SetSystemCursor.restype  = ctypes.wintypes.BOOL
user32.SetSystemCursor.argtypes = [ctypes.c_void_p, ctypes.wintypes.DWORD]
user32.CreateCursor.restype  = ctypes.c_void_p
user32.CreateCursor.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int,
                                 ctypes.c_int, ctypes.c_int,
                                 ctypes.c_void_p, ctypes.c_void_p]

def _hide_cursor():
    """모든 시스템 커서를 투명 커서로 교체"""
    hblank = _make_transparent_cursor()
    if not hblank:
        return
    for cid in _CURSOR_IDS:
        try:
            # 원본 백업 (CopyCursor로 복사)
            orig = user32.LoadCursorW(None, ctypes.c_void_p(cid))
            if orig and cid not in _original_cursors:
                _original_cursors[cid] = user32.CopyIcon(orig)
            # 투명으로 교체 (SetSystemCursor는 핸들을 소유하므로 매번 새로 생성)
            blank = _make_transparent_cursor()
            if blank:
                user32.SetSystemCursor(blank, cid)
        except Exception:
            pass
    log_info("[CURSOR] 커서 숨김 (투명 교체)")

def _restore_cursor():
    """원본 커서 복원"""
    try:
        user32.SystemParametersInfoW(0x0057, 0, None, 0)  # SPI_SETCURSORS
        log_info("[CURSOR] 커서 복원 (SystemParametersInfo)")
    except Exception as e:
        log_info(f"[WARN] 커서 복원 실패: {e}")

# ──────────────────────────────────────────────
# 입력 차단 ON/OFF
# ──────────────────────────────────────────────
def set_block(block: bool):
    global _block_input
    if _block_input != block:
        _block_input = block
        log_info(f"[BLOCK] 입력 차단: {'ON' if block else 'OFF'}")
        if block:
            _hide_cursor()
        else:
            _restore_cursor()

# ──────────────────────────────────────────────
# PC 전환 (훅 콜백에서 호출 - 별도 스레드로 실행)
# ──────────────────────────────────────────────
# 세컨드 PC의 현재 마우스 논리 좌표 추적
_second_cur_x = 0
_second_cur_y = 0
# 메인 PC 마우스 주차 위치 (세컨드 제어 중 이 위치로 계속 복귀)
_park_x = 0
_park_y = 0
# 세컨드 PC 물리 해상도 (DPI 스케일 적용 전) - 마우스 델타 계산용
_second_phys_w = 1920
_second_phys_h = 1080
# 패킷 레이트 카운터
_pkt_sent_count  = 0
_pkt_sent_lock   = threading.Lock()

def _do_switch(target: str, cur_x: int, cur_y: int):
    global active_pc, switching, _second_cur_x, _second_cur_y, _park_x, _park_y
    if switching or active_pc == target:
        return
    switching = True
    log_info(f"[SWITCH] {active_pc} --> {target}  at ({cur_x},{cur_y})")

    if CLIPBOARD_SYNC:
        if target == "second":
            _push_clipboard()
        else:
            _pull_clipboard()
            time.sleep(0.12)

    active_pc = target
    if target == "second":
        set_block(True)
        # 메인 마우스를 복귀 트리거 쪽에 주차
        px, py = park_for_second(cur_x, cur_y)
        ipx, ipy = int(px), int(py)
        pyautogui.moveTo(ipx, ipy)
        _park_x, _park_y = ipx, ipy
        log_info(f"[PARK] 주차위치 설정: ({ipx},{ipy})")
        # 세컨드 커서를 진입 위치(연속 화면)로 이동
        sw, sh = _second_screen_w, _second_screen_h
        ex, ey = entry_for_second(cur_x, cur_y, sw, sh)
        ex, ey = max(0, min(sw-1, ex)), max(0, min(sh-1, ey))
        _second_cur_x, _second_cur_y = ex, ey
        # 진입 이벤트는 warp (즉시 이동, 보간 없음)
        send_event({'type': 'mouse_warp', 'x': ex, 'y': ey})
        log_info(f"[SWITCH] 세컨드 시작. 메인주차:({ipx},{ipy}) 세컨드진입:({ex},{ey})")
    else:
        set_block(False)
        # 메인 마우스를 세컨드와 맞닿은 쪽 끝, y는 세컨드 현재 y 비율로
        sw, sh = _second_screen_w, _second_screen_h
        mx, my = return_for_main(_second_cur_x, _second_cur_y, sw, sh)
        mx = max(VS_LEFT, min(VS_RIGHT-1, mx))
        my = max(VS_TOP,  min(VS_BOTTOM-1, my))
        pyautogui.moveTo(int(mx), int(my))
        log_info(f"[SWITCH] 메인 복귀. 위치:({int(mx)},{int(my)})")

    time.sleep(0.08)
    switching = False

def switch_to(target, x, y):
    threading.Thread(target=_do_switch, args=(target, x, y), daemon=True).start()

# ──────────────────────────────────────────────
# Windows 훅 콜백
# ──────────────────────────────────────────────
def _mouse_hook_proc(nCode, wParam, lParam):
    global _second_cur_x, _second_cur_y, _pkt_sent_count
    if nCode == HC_ACTION:
        try:
            ms  = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT))[0]
            x, y = ms.pt.x, ms.pt.y
            injected = bool(ms.flags & LLMHF_INJECTED)

            if wParam == WM_MOUSEMOVE and not injected:
                if active_pc == "main":
                    if to_second_cond(x, y):
                        log(f"[EDGE] ({x},{y}) --> 세컨드 전환")
                        switch_to("second", x, y)
                else:
                    # 세컨드 제어 중:
                    # 1) 실제 물리 델타를 세컨드 좌표로 변환해서 전달
                    # 2) 메인 마우스는 주차위치로 되돌림 (SetCursorPos = injected)
                    dx = x - _park_x
                    dy = y - _park_y
                    if dx != 0 or dy != 0:
                        # 델타를 그대로 논리 좌표로 사용
                        # 근거: LL훅의 pt.x/y는 OS 가상 데스크톱 기준 논리 픽셀
                        # 메인/세컨드 DPI 무관하게 dx/dy를 그대로 세컨드 논리 좌표에 더하면
                        # pyautogui(논리 픽셀 기준)가 정확히 그만큼 이동함
                        tx = max(0, min(_second_screen_w-1, _second_cur_x + dx))
                        ty = max(0, min(_second_screen_h-1, _second_cur_y + dy))
                        ev = {'type': 'mouse_move_abs', 'x': tx, 'y': ty}
                        # 중간 패킷 버리기 옵션
                        if MOUSE_SKIP_PACKETS:
                            try:
                                tmp = []
                                while not _send_queue.empty():
                                    try:
                                        item = _send_queue.get_nowait()
                                        if item.get('type') != 'mouse_move_abs':
                                            tmp.append(item)
                                    except _queue.Empty:
                                        break
                                for item in tmp:
                                    _send_queue.put_nowait(item)
                            except Exception:
                                pass
                        try:
                            _send_queue.put_nowait(ev)
                        except _queue.Full:
                            pass
                        # 패킷 레이트 카운터 증가
                        _pkt_sent_count += 1
                        _second_cur_x, _second_cur_y = tx, ty
                        if LOG_MOUSE_MOVE:
                            log(f"[MOVE] delta({dx},{dy}) -> second({tx},{ty})")
                        # 복귀 엣지: 세컨드 커서가 세컨드 화면 끝에 도달했을 때
                        if to_main_cond_second(tx, ty):
                            log(f"[EDGE] second({tx},{ty}) --> 메인 복귀")
                            switch_to("main", x, y)
                    # 메인 마우스를 주차위치로 되돌림 (이벤트는 차단)
                    user32.SetCursorPos(_park_x, _park_y)
                    return 1

            elif not injected and wParam != WM_MOUSEMOVE:
                if active_pc == "second":
                    btn_map = {
                        WM_LBUTTONDOWN: ('Button.left',   True),
                        WM_LBUTTONUP:   ('Button.left',   False),
                        WM_RBUTTONDOWN: ('Button.right',  True),
                        WM_RBUTTONUP:   ('Button.right',  False),
                        WM_MBUTTONDOWN: ('Button.middle', True),
                        WM_MBUTTONUP:   ('Button.middle', False),
                    }
                    if wParam in btn_map:
                        btn, pressed = btn_map[wParam]
                        send_event({'type': 'mouse_click', 'button': btn, 'pressed': pressed})
                        log(f"[CLICK] {btn} {'down' if pressed else 'up'} -> second")
                    elif wParam == WM_MOUSEWHEEL:
                        delta = ctypes.c_short(ms.mouseData >> 16).value
                        dy = 1 if delta > 0 else -1
                        send_event({'type': 'mouse_scroll', 'dx': 0, 'dy': dy})
                        log(f"[SCROLL] dy={dy} -> second")
                    # 차단
                    return 1
        except Exception as e:
            log(f"[ERR] mouse hook: {e}")

    return user32.CallNextHookEx(_mouse_hook_id, nCode, wParam, lParam)

# keybd_event 플래그
KEYEVENTF_KEYUP       = 0x0002
KEYEVENTF_EXTENDEDKEY = 0x0001
# 우리가 주입한 이벤트 구분용 마커 (dwExtraInfo에 설정)
KVM_MARKER = 0x4B564D31  # ASCII 'KVM1'

# SendInput 구조체 정의 (dwExtraInfo 확실히 전달하기 위해 keybd_event 대신 사용)
INPUT_KEYBOARD = 1

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk",         ctypes.wintypes.WORD),
        ("wScan",       ctypes.wintypes.WORD),
        ("dwFlags",     ctypes.wintypes.DWORD),
        ("time",        ctypes.wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]

class _INPUT_UNION(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("_pad", ctypes.c_byte * 32)]

class INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.wintypes.DWORD), ("_u", _INPUT_UNION)]

user32.SendInput.restype  = ctypes.wintypes.UINT
user32.SendInput.argtypes = [ctypes.wintypes.UINT,
                              ctypes.POINTER(INPUT),
                              ctypes.c_int]

def _send_key_event(vk: int, flags: int):
    """KVM_MARKER가 설정된 키 이벤트 주입 (우리 훅이 무시하도록)"""
    inp = INPUT()
    inp.type = INPUT_KEYBOARD
    inp._u.ki.wVk         = vk
    inp._u.ki.wScan       = 0
    inp._u.ki.dwFlags     = flags
    inp._u.ki.time        = 0
    inp._u.ki.dwExtraInfo = KVM_MARKER
    user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

def _kb_hook_proc(nCode, wParam, lParam):
    if nCode == HC_ACTION:
        try:
            kb = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT))[0]
            is_dn = wParam in (WM_KEYDOWN, WM_SYSKEYDOWN)

            # ── 우리가 주입한 이벤트 판별: dwExtraInfo 마커로 확인 ──
            # flags의 LLMHF_INJECTED 비트는 Ctrl/Alt 등 실제 키도 1로 올 수 있어 신뢰 불가
            # 대신 keybd_event 호출 시 dwExtraInfo=KVM_MARKER를 설정해서 구분
            our_event = (kb.dwExtraInfo == KVM_MARKER)

            if not our_event:
                # ── 비상 탈출: 최우선, 항상 감지 ──
                if is_dn:
                    _emergency_keys_down.add(kb.vkCode)
                else:
                    _emergency_keys_down.discard(kb.vkCode)

                if EMERGENCY_COMBO.issubset(_emergency_keys_down) and active_pc == "second":
                    log_info("[EMERGENCY] ★★★ Ctrl+Alt+F12 감지! 강제 복귀 ★★★")
                    _emergency_keys_down.clear()
                    threading.Thread(target=_emergency_release, daemon=True).start()
                    return 1

                # ── KEY-RAW 진단 로그 (verbose 모드) ──
                log(f"[KEY-RAW] vk={kb.vkCode:#04x}({kb.vkCode}) "
                    f"wParam={wParam:#06x} flags={kb.flags:#010x} "
                    f"{'DN' if is_dn else 'UP'} active={active_pc}")

                # ── 세컨드 PC 제어 중: 전달 + 차단 ──
                if active_pc == "second":
                    # vk 숫자를 그대로 전송 → 클라이언트가 SendInput으로 정확히 재생
                    # extended key 여부도 함께 전송
                    ext = 1 if kb.vkCode in _EXTENDED_KEYS else 0
                    etype = 'key_dn' if is_dn else 'key_up'
                    send_event({'type': etype, 'vk': kb.vkCode,
                                'scan': kb.scanCode, 'ext': ext})
                    log(f"[KEY] {'DN' if is_dn else 'UP'} vk={kb.vkCode:#04x} ext={ext}")
                    # 윈도우키: 즉시 키업 주입으로 시작메뉴 차단
                    if kb.vkCode in _WIN_KEYS and is_dn:
                        _send_key_event(kb.vkCode, KEYEVENTF_KEYUP)
                    return 1  # 메인 PC에 전달 차단

        except Exception as e:
            log_info(f"[ERR] kb hook: {e}")

    return user32.CallNextHookEx(_kb_hook_id, nCode, wParam, lParam)

_mouse_hook_cb = HOOKPROC(_mouse_hook_proc)
_kb_hook_cb    = HOOKPROC(_kb_hook_proc)

# ──────────────────────────────────────────────
# 메시지 루프 (메인 스레드에서 실행 필수)
# ──────────────────────────────────────────────
def run_hook_loop():
    """훅 설치 + GetMessage 루프 - 반드시 메인 스레드에서 호출"""
    global _kb_hook_id, _mouse_hook_id, _hook_thread_id

    _hook_thread_id = kernel32.GetCurrentThreadId()
    log_info(f"[HOOK] 스레드 ID: {_hook_thread_id}")

    # Low-level 훅은 hmod=NULL, dwThreadId=0 (전역)
    _mouse_hook_id = user32.SetWindowsHookExW(WH_MOUSE_LL,    _mouse_hook_cb, 0, 0)
    _kb_hook_id    = user32.SetWindowsHookExW(WH_KEYBOARD_LL, _kb_hook_cb,    0, 0)

    if _mouse_hook_id and _kb_hook_id:
        log_info(f"[HOOK] 설치 완료 (mouse={_mouse_hook_id}, kb={_kb_hook_id})")
    else:
        err = kernel32.GetLastError()
        log_info(f"[ERR] 훅 설치 실패 LastError={err}")
        return

    log_info("[READY] 준비 완료! 마우스를 화면 끝으로 이동하면 전환됩니다.")

    # GetMessage 루프 - 훅이 살아있으려면 이 루프가 돌아야 함
    msg = ctypes.wintypes.MSG()
    while True:
        ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
        if ret == 0 or ret == -1:
            break
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

    if _kb_hook_id:    user32.UnhookWindowsHookEx(_kb_hook_id)
    if _mouse_hook_id: user32.UnhookWindowsHookEx(_mouse_hook_id)
    log_info("[HOOK] 훅 해제 완료")

# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────
def main():
    if SECOND_PC_IP == "192.168.0.XXX":
        print("[ERR] kvm_config.ini 에서 second_pc_ip 를 변경하세요!")
        input("Enter 키를 눌러 종료...")
        sys.exit(1)

    dirs = {
        "left":   ("왼쪽 끝",  "오른쪽 끝"),
        "right":  ("오른쪽 끝","왼쪽 끝"),
        "top":    ("위쪽 끝",  "아래쪽 끝"),
        "bottom": ("아래쪽 끝","위쪽 끝"),
    }
    to_lbl, back_lbl = dirs.get(SECOND_PC_POSITION, ("왼쪽 끝","오른쪽 끝"))

    print("=" * 65)
    print("  Software KVM - 메인 PC v5.0  (pynput 제거, 순수 WinAPI)")
    print("=" * 65)
    print(f"  설정 파일       : {CONFIG_FILE}")
    print(f"  세컨드 PC IP    : {SECOND_PC_IP}:{PORT}")
    print(f"  세컨드 PC 방향  : {SECOND_PC_POSITION.upper()}")
    print(f"  가상 화면       : {VS_W}x{VS_H}  (오프셋 {VS_LEFT},{VS_TOP})")
    if MONITORS:
        print(f"  감지된 모니터   : {len(MONITORS)}개")
        for i, m in enumerate(MONITORS):
            print(f"    [{i+1}] {m.width}x{m.height} @ ({m.x},{m.y})")
    print(f"  클립보드 동기화 : {'ON' if CLIPBOARD_SYNC else 'OFF'}")
    print(f"  상세 로그       : {'ON' if VERBOSE else 'OFF'}")
    print(f"  엣지 감지 범위  : {EDGE_THRESHOLD}px")
    print()
    print(f"  [전환] 마우스를 화면 {to_lbl}   --> 세컨드 PC")
    print(f"  [복귀] 마우스를 화면 {back_lbl} --> 메인 PC")
    print()
    print("  세컨드 PC 연결 시도 중... (종료: Ctrl+C 또는 창 닫기)")
    print()
    print("  [비상 탈출] Ctrl+Alt+F12 --> 강제 메인 복귀 (차단 해제)")
    print("=" * 65)

    # 네트워크 스레드 (백그라운드)
    threading.Thread(target=_send_worker,        daemon=True).start()
    threading.Thread(target=_packet_rate_logger, daemon=True).start()
    threading.Thread(target=connect_loop,        daemon=True).start()
    threading.Thread(target=recv_loop,           daemon=True).start()

    # Ctrl+C 처리 (메인 스레드가 GetMessage에 있으므로 별도 처리)
    def _ctrl_c_watcher():
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            log_info("[EXIT] Ctrl+C 감지, 종료 중...")
            if _hook_thread_id:
                user32.PostThreadMessageW(_hook_thread_id, WM_QUIT, 0, 0)
    threading.Thread(target=_ctrl_c_watcher, daemon=True).start()

    # 메인 스레드에서 훅 + GetMessage 루프 실행
    try:
        run_hook_loop()
    except KeyboardInterrupt:
        log_info("[EXIT] 종료")

if __name__ == '__main__':
    main()
