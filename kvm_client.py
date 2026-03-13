"""
==================================================
  Software KVM - 세컨드 PC (수신 측) v4.0
==================================================
설치: pip install pynput pyautogui pyperclip
설정: kvm_config.ini 의 port 확인
실행: python kvm_client.py
==================================================
"""

import socket
import json
import threading
import time
import os
import sys
import configparser
import ctypes
import ctypes.wintypes
import pyautogui
import pyperclip
from pynput.mouse import Controller as MouseCtrl, Button

pyautogui.FAILSAFE = False
pyautogui.PAUSE    = 0

# ──────────────────────────────────────────────
# 설정 파일 로드
# ──────────────────────────────────────────────
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kvm_config.ini")

def load_config():
    cfg = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        cfg.read(CONFIG_FILE, encoding='utf-8')
    return cfg

cfg            = load_config()
PORT           = cfg.getint     ('network', 'port',           fallback=9999)
VERBOSE          = cfg.getboolean ('log',      'verbose',          fallback=True)
LOG_MOUSE_MOVE   = cfg.getboolean ('log',      'log_mouse_move',    fallback=False)
LOG_PACKET_RATE  = cfg.getboolean ('log',      'log_packet_rate',   fallback=False)
# smooth_move: 목표 좌표까지 보간하여 부드럽게 이동
SMOOTH_MOVE      = cfg.getboolean ('features', 'smooth_move',       fallback=True)
# smooth_interval_ms: 보간 스텝 간격 (ms). 낮을수록 부드럽고 CPU 사용 증가
SMOOTH_INTERVAL  = cfg.getfloat   ('features', 'smooth_interval_ms',fallback=4.0)
# smooth_steps: 한 번 이동 시 보간 스텝 수. 낮으면 빠르고 높으면 더 부드러움
SMOOTH_STEPS     = cfg.getint     ('features', 'smooth_steps',      fallback=3)

# ──────────────────────────────────────────────
# 로깅
# ──────────────────────────────────────────────
_log_lock = threading.Lock()

def log(msg: str):
    if VERBOSE:
        with _log_lock:
            print(f"[{time.strftime('%H:%M:%S.') + f'{int(time.time()*1000)%1000:03d}'}] {msg}", flush=True)

def log_info(msg: str):
    with _log_lock:
        print(f"[{time.strftime('%H:%M:%S.') + f'{int(time.time()*1000)%1000:03d}'}] {msg}", flush=True)

# ──────────────────────────────────────────────
# DPI 설정 (정확한 좌표 처리)
# ──────────────────────────────────────────────
try:
    # Per-Monitor DPI Aware v2 (Windows 10+)
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    log_info("[DPI] SetProcessDpiAwareness(2) 설정")
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
        log_info("[DPI] SetProcessDPIAware() 설정 (fallback)")
    except Exception:
        pass

def get_dpi_scale():
    try:
        dpi = ctypes.windll.user32.GetDpiForSystem()
        return dpi / 96.0
    except Exception:
        try:
            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
            ctypes.windll.user32.ReleaseDC(0, hdc)
            return dpi / 96.0
        except Exception:
            return 1.0

DPI_SCALE         = get_dpi_scale()
SCREEN_W, SCREEN_H = pyautogui.size()

mouse = MouseCtrl()

# ──────────────────────────────────────────────
# SendInput으로 키보드 이벤트 재생 (pynput 대체)
# pynput은 특수문자/IME 키를 정확히 재생하지 못함
# SendInput + vk코드 방식은 OS 수준에서 동일하게 동작
# ──────────────────────────────────────────────
import struct as _struct

KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP       = 0x0002

# Windows INPUT 구조체는 정확히 28바이트
# ctypes 구조체로 만들면 64bit Python에서 패딩으로 40바이트가 되어
# SendInput(nSize=40)이 실패함 → raw 바이트 버퍼로 직접 구성
#
# INPUT layout (28 bytes):
#  [0] type       DWORD  4  = 1 (INPUT_KEYBOARD)
#  [4] wVk        WORD   2
#  [6] wScan      WORD   2
#  [8] dwFlags    DWORD  4
# [12] time       DWORD  4
# [16] dwExtraInfo UINT64 8  (x64) / UINT32 4 (x86)
# total = 28 (x86) or 32? 
# 실제 Windows SDK: sizeof(INPUT)=28 (x86), sizeof(INPUT)=40 (x64) 아님
# MSDN: sizeof(INPUT) = 28 on both x86 and x64
# → 28바이트 고정 버퍼 사용

# keybd_event 사용 - SendInput보다 단순하고 ARM/x64/x86 모두 동작
# keybd_event(bVk, bScan, dwFlags, dwExtraInfo)
# 구조체 없이 4개 인자만 받으므로 플랫폼 레이아웃 문제 없음

_user32 = ctypes.windll.user32
_user32.keybd_event.restype  = None
_user32.keybd_event.argtypes = [
    ctypes.wintypes.BYTE,   # bVk
    ctypes.wintypes.BYTE,   # bScan
    ctypes.wintypes.DWORD,  # dwFlags
    ctypes.POINTER(ctypes.wintypes.ULONG),  # dwExtraInfo
]

def replay_key(vk: int, scan: int, ext: int, is_down: bool):
    """vk코드로 키 이벤트 재생 - ARM/x64/x86 모든 플랫폼 동작"""
    flags = 0
    if ext:
        flags |= KEYEVENTF_EXTENDEDKEY
    if not is_down:
        flags |= KEYEVENTF_KEYUP
    _user32.keybd_event(vk & 0xFF, scan & 0xFF, flags, None)

# 수신 패킷 카운터
_pkt_recv_count = 0
_pkt_recv_lock  = threading.Lock()

# ──────────────────────────────────────────────
# 부드러운 마우스 이동 엔진
# ──────────────────────────────────────────────
# 서버로부터 목표 좌표(tx, ty)를 받으면
# 별도 스레드가 현재 위치 → 목표 위치를 SMOOTH_STEPS 단계로 보간
# 새 목표가 오면 즉시 교체(현재 진행 중인 보간 중단)

_smooth_target   = None          # (tx, ty) 최신 목표
_smooth_lock     = threading.Lock()
_smooth_event    = threading.Event()  # 새 목표 신호
_cur_x           = SCREEN_W // 2     # 클라이언트 현재 논리 위치 추적
_cur_y           = SCREEN_H // 2

def _smooth_worker():
    """목표 좌표까지 보간 이동하는 전용 스레드"""
    global _cur_x, _cur_y, _smooth_target
    interval = SMOOTH_INTERVAL / 1000.0   # ms → sec

    while True:
        _smooth_event.wait()              # 새 목표 도착까지 대기
        _smooth_event.clear()

        while True:
            with _smooth_lock:
                target = _smooth_target

            if target is None:
                break

            tx, ty = target
            cx, cy = _cur_x, _cur_y
            dx, dy = tx - cx, ty - cy

            if abs(dx) < 2 and abs(dy) < 2:
                # 목표에 거의 도달 → 정확히 이동 후 종료
                if abs(dx) > 0 or abs(dy) > 0:
                    mouse.position = (tx, ty)
                    _cur_x, _cur_y = tx, ty
                with _smooth_lock:
                    _smooth_target = None
                break

            # SMOOTH_STEPS 단계로 나눠 이동
            steps = max(1, min(SMOOTH_STEPS, max(abs(dx), abs(dy)) // 8))
            nx = cx + dx // steps
            ny = cy + dy // steps

            mouse.position = (nx, ny)
            _cur_x, _cur_y = nx, ny
            time.sleep(interval)

def set_mouse_target(tx: int, ty: int):
    """목표 좌표 설정 - 보간 스레드가 부드럽게 이동시킴"""
    global _smooth_target
    with _smooth_lock:
        _smooth_target = (tx, ty)
    _smooth_event.set()

def _packet_rate_logger():
    global _pkt_recv_count
    while True:
        time.sleep(1.0)
        if LOG_PACKET_RATE:
            with _pkt_recv_lock:
                count = _pkt_recv_count
                _pkt_recv_count = 0
            if count > 0:
                log_info(f"[PKT-RATE] 수신: {count} pkt/s")
        else:
            with _pkt_recv_lock:
                _pkt_recv_count = 0

BUTTON_MAP = {
    'Button.left':   Button.left,
    'Button.right':  Button.right,
    'Button.middle': Button.middle,
}

# ──────────────────────────────────────────────
# 이벤트 처리
# ──────────────────────────────────────────────
_last_clip = ""

def handle_event(event: dict, conn: socket.socket):
    global _last_clip
    etype = event.get('type')

    if etype == 'mouse_warp':
        # switch 진입/복귀 시 즉시 이동 (보간 없음)
        x, y = event['x'], event['y']
        mouse.position = (x, y)
        # 보간 스레드의 현재 위치도 동기화
        global _cur_x, _cur_y
        _cur_x, _cur_y = x, y
        with _smooth_lock:
            _smooth_target = None  # 진행 중인 보간 취소
        if LOG_MOUSE_MOVE:
            log(f"[WARP] -> ({x},{y})")

    elif etype == 'mouse_move_abs':
        global _pkt_recv_count
        x, y = event['x'], event['y']
        if SMOOTH_MOVE:
            set_mouse_target(x, y)
        else:
            mouse.position = (x, y)
        with _pkt_recv_lock:
            _pkt_recv_count += 1
        if LOG_MOUSE_MOVE:
            log(f"[MOVE] -> ({x},{y})")

    elif etype == 'mouse_move':
        x = int(event['rx'] * SCREEN_W)
        y = int(event['ry'] * SCREEN_H)
        if SMOOTH_MOVE:
            set_mouse_target(x, y)
        else:
            mouse.position = (x, y)
        if LOG_MOUSE_MOVE:
            log(f"[MOVE-rel] -> ({x},{y})")

    elif etype == 'mouse_click':
        btn = BUTTON_MAP.get(event['button'], Button.left)
        if event['pressed']:
            mouse.press(btn)
        else:
            mouse.release(btn)
        log(f"[CLICK] {event['button']} {'down' if event['pressed'] else 'up'}")

    elif etype == 'mouse_scroll':
        mouse.scroll(event['dx'], event['dy'])
        log(f"[SCROLL] ({event['dx']},{event['dy']})")

    elif etype == 'key_dn':
        vk   = event.get('vk', 0)
        scan = event.get('scan', 0)
        ext  = event.get('ext', 0)
        replay_key(vk, scan, ext, True)
        log(f"[KEY] DN vk={vk:#04x}")

    elif etype == 'key_up':
        vk   = event.get('vk', 0)
        scan = event.get('scan', 0)
        ext  = event.get('ext', 0)
        replay_key(vk, scan, ext, False)
        log(f"[KEY] UP vk={vk:#04x}")

    # 구버전 호환 (key_press/key_release)
    elif etype == 'key_press':
        vk = event.get('vk', 0)
        if vk:
            replay_key(vk, 0, event.get('ext', 0), True)
        log(f"[KEY] DN(legacy) vk={vk:#04x}")

    elif etype == 'key_release':
        vk = event.get('vk', 0)
        if vk:
            replay_key(vk, 0, event.get('ext', 0), False)
        log(f"[KEY] UP(legacy) vk={vk:#04x}")

    elif etype == 'clipboard':
        text = event.get('text', '')
        if text and text != _last_clip:
            _last_clip = text
            try:
                pyperclip.copy(text)
                log_info(f"[CB] <- 메인: {text[:60]}")
            except Exception as e:
                log_info(f"[WARN] 클립보드 실패: {e}")

    elif etype == 'clipboard_request':
        try:
            text = pyperclip.paste()
            if text and text != _last_clip:
                _last_clip = text
                conn.sendall((json.dumps({'type': 'clipboard', 'text': text}) + '\n').encode('utf-8'))
                log_info(f"[CB] -> 메인: {text[:60]}")
        except Exception as e:
            log_info(f"[WARN] 클립보드 읽기 실패: {e}")

    elif etype == 'dpi_request':
        # SCREEN_W/H = pyautogui.size() = 논리 해상도
        # 물리 해상도 = 논리 * DPI_SCALE
        phys_w = int(SCREEN_W * DPI_SCALE)
        phys_h = int(SCREEN_H * DPI_SCALE)
        info = {
            'type':   'dpi_info',
            'scale':  DPI_SCALE,
            'width':  phys_w,    # 물리 해상도 전송
            'height': phys_h,
            'logical_w': SCREEN_W,   # 논리 해상도도 함께 전송
            'logical_h': SCREEN_H,
        }
        try:
            conn.sendall((json.dumps(info) + '\n').encode('utf-8'))
            log_info(f"[DPI] 정보 전송: 물리={phys_w}x{phys_h}, 논리={SCREEN_W}x{SCREEN_H}, scale={DPI_SCALE:.2f}x")
        except Exception as e:
            log_info(f"[WARN] DPI 전송 실패: {e}")

# ──────────────────────────────────────────────
# 연결 처리
# ──────────────────────────────────────────────
def handle_connection(conn: socket.socket, addr):
    log_info(f"[OK] 메인 PC 연결됨: {addr[0]}")
    buf = ""
    try:
        while True:
            data = conn.recv(4096).decode('utf-8', errors='ignore')
            if not data:
                break
            buf += data
            while '\n' in buf:
                line, buf = buf.split('\n', 1)
                line = line.strip()
                if line:
                    try:
                        handle_event(json.loads(line), conn)
                    except json.JSONDecodeError as e:
                        log(f"[WARN] JSON 오류: {e} | {line[:80]}")
    except (ConnectionResetError, BrokenPipeError, OSError) as e:
        log_info(f"[WARN] 연결 오류: {e}")
    finally:
        conn.close()
        log_info(f"[--] 연결 해제: {addr[0]}")
        log_info("[..] 메인 PC 재연결 대기 중...")

# ──────────────────────────────────────────────
# 로컬 IP
# ──────────────────────────────────────────────
def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"
    finally:
        s.close()

# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────
def main():
    local_ip = get_local_ip()

    print("=" * 60)
    print("  Software KVM - 세컨드 PC v4.0")
    print("=" * 60)
    print(f"  이 PC의 IP 주소 : {local_ip}")
    print(f"  포트            : {PORT}")
    print(f"  화면 해상도     : {SCREEN_W} x {SCREEN_H}")
    print(f"  DPI 스케일      : {DPI_SCALE:.2f}x  ({int(DPI_SCALE*100)}%)")
    print(f"  상세 로그       : {'ON' if VERBOSE else 'OFF'}")
    print()
    print("  [설정] 메인 PC의 kvm_config.ini 에 아래 값을 입력하세요:")
    print(f"         second_pc_ip = {local_ip}")
    print()
    print("  메인 PC 연결 대기 중... (종료: Ctrl+C)")
    print("=" * 60)

    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind(('0.0.0.0', PORT))
    server.listen(1)
    threading.Thread(target=_packet_rate_logger, daemon=True).start()
    if SMOOTH_MOVE:
        threading.Thread(target=_smooth_worker, daemon=True).start()
        log_info(f"[SMOOTH] 보간 이동 ON  (steps={SMOOTH_STEPS}, interval={SMOOTH_INTERVAL}ms)")
    log_info(f"[LISTEN] 포트 {PORT} 대기 중...")

    try:
        while True:
            conn, addr = server.accept()
            threading.Thread(target=handle_connection, args=(conn, addr), daemon=True).start()
    except KeyboardInterrupt:
        print("\n종료 중...")
    finally:
        server.close()

if __name__ == '__main__':
    main()
