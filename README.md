# Simple Software KVM

두 대의 Windows PC를 하나의 키보드/마우스로 제어합니다.

---

## 설치

```bash
pip install pynput pyautogui pyperclip screeninfo
```

---

## 파일 구성

```
kvm_server.py     ← 메인 PC에서 실행
kvm_client.py     ← 세컨드 PC에서 실행
kvm_config.ini    ← 설정 파일 (양쪽 모두 같은 폴더에)
```

---

## 설정 (kvm_config.ini)

```ini
[network]
second_pc_ip = 192.168.0.101   # 세컨드 PC IP (필수 변경)
port = 9999
reconnect_sec = 3

[layout]
second_pc_position = left      # left | right | top | bottom
edge_threshold = 8             # 클수록 전환이 쉬워짐

[features]
clipboard_sync = true

[log]
verbose = false                # true: 모든 이벤트 로그 출력
log_mouse_move = false         # true: 마우스 이동 로그도 출력 (매우 많음)
```

---

## 실행 순서

**1. 세컨드 PC 먼저 실행**
```
python kvm_client.py
```
출력된 IP를 메모 → `kvm_config.ini`의 `second_pc_ip`에 입력

**2. 메인 PC 실행**
```
python kvm_server.py
```

---

## 전환 방법

`second_pc_position = left` 기준:

| 동작 | 결과 |
|------|------|
| 마우스를 화면 **왼쪽 끝** 이동 | 세컨드 PC 제어 시작 |
| 마우스를 화면 **오른쪽 끝** 이동 | 메인 PC 복귀 |

---

## Windows 방화벽 설정 (세컨드 PC, 관리자 PowerShell)

```powershell
netsh advfirewall firewall add rule name="KVM Port 9999" dir=in action=allow protocol=TCP localport=9999
```

---

## 문제 해결

| 증상 | 해결책 |
|------|--------|
| 연결 안 됨 | 방화벽 포트 9999 허용 확인 |
| 관리자 앱 제어 안 됨 | Python을 관리자 권한으로 실행 |
| 마우스 위치 어긋남 | 세컨드 PC 실행 후 DPI 스케일 확인 로그 확인 |
| 전환이 너무 쉬움 | `edge_threshold = 3` 으로 줄이기 |
| 로그가 너무 많음 | `verbose = false`, `log_mouse_move = false` |
