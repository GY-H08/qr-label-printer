# QR Label Printer

폼텍 라벨지에 QR코드가 포함된 박스 라벨을 자동 출력하는 Windows 데스크탑 프로그램.

## 주요 기능
- 폼텍 라벨지 2종 지원 (3102: 40칸, 3104: 27칸)
- QR코드 자동 생성 및 출력
- 라벨 ID 자동 생성 (소재지/모듈/날짜/일련번호 기반)
- config.json으로 소재지·모듈 코드 중앙 관리
- 출력 회차 및 일련번호 캐시로 중복 방지
- PyQt5 기반 GUI

## 라벨 ID 구조
KR-[소재지코드][모듈코드][task index]-[날짜YYMMDD]-[연번 000~999]
예시: KR-AA1A1001-260303-001

## 기술 스택
- Python, PyQt5
- Windows GDI API (win32print, win32ui)
- qrcode, Pillow

## 실행 방법
```bash
pip install PyQt5 pywin32 qrcode pillow
python qr_label_printer.py
```

## 설정 (setting/config.json)
```json
{
    "site": "소재지코드",
    "module": "모듈코드"
}
```
