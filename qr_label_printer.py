import sys
import os
import json
from datetime import datetime

import PyQt5

os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = os.path.join(
    os.path.dirname(PyQt5.__file__),
    "Qt5", "plugins", "platforms"
)

import win32print
import win32ui
import win32con
from PIL import ImageWin, Image
import qrcode

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton,
    QVBoxLayout, QComboBox, QLineEdit,
    QMessageBox, QGroupBox
)





def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)  # exe 위치
    else:
        return os.path.dirname(os.path.abspath(__file__))  # py 위치

BASE_DIR = get_base_dir()
SETTING_DIR = os.path.join(BASE_DIR, "setting")
CACHE_FILE  = os.path.join(SETTING_DIR, "print_cache_2.json")
CONFIG_FILE = os.path.join(SETTING_DIR, "config.json")

A4_W_MM = 210.0
A4_H_MM = 297.0





def load_config():
    if not os.path.exists(SETTING_DIR):
        os.makedirs(SETTING_DIR)

    if not os.path.exists(CONFIG_FILE):
        default = {"site": "SITE_CODE", "module": "MODULE_CODE"}
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(default, f, indent=4, ensure_ascii=False)
        return default

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)




LABEL_SPECS = {

    "40칸 - 3102 (47×26.9mm)": {
        "label_w": 47.0,  "label_h": 26.9,
        "top":     14.2,  "left":     7.0,
        "col_gap": 49.5,  "row_gap": 27.5,
        "cols":     4,    "rows":    10,
        "offset_x": 0.0,  "offset_y": -5.0,
    },

    "27칸 - 3104 (62.7×30.1mm)": {
        "label_w": 62.7,  "label_h": 30.1,
        "top":     11.0,  "left":     8.4,
        "col_gap": 65.2,  "row_gap": 30.1,
        "cols":     3,    "rows":     9,
        "offset_x": 0.0,  "offset_y": 0.0,
    },
}





def load_cache():
    if not os.path.exists(SETTING_DIR):
        os.makedirs(SETTING_DIR)

    today = datetime.now().strftime("%y%m%d")

    if not os.path.exists(CACHE_FILE):
        data = {"date": today, "daily_task_index": 0, "print_round": 0, "serial": {}}
        save_cache(data)
        return data

    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    if data.get("date") != today:
        data = {"date": today, "daily_task_index": 0, "print_round": 0, "serial": {}}
        save_cache(data)

    return data


def save_cache(data):
    with open(CACHE_FILE, "w") as f:
        json.dump(data, f, indent=4)


def mm_to_px(mm, total_px, total_mm):
    return int(round(mm * total_px / total_mm))





def print_gdi(labels, printer_name, spec):
    per_page = spec["cols"] * spec["rows"]

    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)

        page_w_px = hDC.GetDeviceCaps(win32con.HORZRES)
        page_h_px = hDC.GetDeviceCaps(win32con.VERTRES)

        top      = mm_to_px(spec["top"],      page_h_px, A4_H_MM)
        left     = mm_to_px(spec["left"],     page_w_px, A4_W_MM)
        label_w  = mm_to_px(spec["label_w"],  page_w_px, A4_W_MM)
        label_h  = mm_to_px(spec["label_h"],  page_h_px, A4_H_MM)
        col_gap  = mm_to_px(spec["col_gap"],  page_w_px, A4_W_MM)
        row_gap  = mm_to_px(spec["row_gap"],  page_h_px, A4_H_MM)
        offset_x = mm_to_px(spec["offset_x"], page_w_px, A4_W_MM)
        offset_y = mm_to_px(spec["offset_y"], page_h_px, A4_H_MM)

        font_mm   = max(1.5, spec["label_h"] * 0.12)
        font_h_px = mm_to_px(font_mm, page_h_px, A4_H_MM)
        font = win32ui.CreateFont({
            "name": "Arial",
            "height": -font_h_px,
            "weight": 400,
        })
        hDC.SelectObject(font)

        text_area_h = int(font_h_px * 1.6)
        margin_top  = int(label_h * 0.04)
        qr_avail_h  = label_h - margin_top - text_area_h
        qr_avail_w  = int(label_w * 0.90)
        qr_size     = max(4, min(qr_avail_h, qr_avail_w))

        hDC.StartDoc("FORMTEC_PRINT")

        for idx, text in enumerate(labels):
            page_index = idx % per_page
            col = page_index % spec["cols"]
            row = page_index // spec["cols"]

            if page_index == 0:
                hDC.StartPage()

            x = left + col * col_gap + offset_x
            y = top  + row * row_gap + offset_y

            qr_x = x + (label_w - qr_size) // 2
            qr_y = y + margin_top + (qr_avail_h - qr_size) // 2

            qr_img = qrcode.make(text).convert("RGB")
            qr_img = qr_img.resize((qr_size, qr_size), Image.NEAREST)
            dib = ImageWin.Dib(qr_img)
            dib.draw(hDC.GetHandleOutput(), (
                qr_x, qr_y,
                qr_x + qr_size,
                qr_y + qr_size
            ))

            hDC.SetTextAlign(win32con.TA_CENTER | win32con.TA_TOP)
            text_x = x + label_w // 2
            text_y = y + label_h - text_area_h
            hDC.TextOut(text_x, text_y, text)
            hDC.SetTextAlign(win32con.TA_LEFT | win32con.TA_TOP)

            if page_index == per_page - 1:
                hDC.EndPage()

        if len(labels) % per_page != 0:
            hDC.EndPage()

        hDC.EndDoc()
        hDC.DeleteDC()

    finally:
        win32print.ClosePrinter(hPrinter)





class LabelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("QR Label Printer")

        config = load_config()
        self.site   = config.get("site", "SITE_CODE")
        self.module = config.get("module", "MODULE_CODE").upper()

        layout = QVBoxLayout()

        label_group = QGroupBox("폼텍 라벨")
        label_inner = QVBoxLayout()
        self.label_combo = QComboBox()
        self.label_combo.addItems(list(LABEL_SPECS.keys()))
        self.label_combo.currentIndexChanged.connect(self.update_info)
        label_inner.addWidget(self.label_combo)
        label_group.setLayout(label_inner)
        layout.addWidget(label_group)

        self.site_label   = QLabel(f"소재지:  {self.site}")
        self.module_label = QLabel(f"모듈 코드:  {self.module}")
        self.site_label.setStyleSheet("font-weight: bold; font-size: 13px; padding: 4px;")
        self.module_label.setStyleSheet("font-weight: bold; font-size: 13px; padding: 4px;")
        layout.addWidget(self.site_label)
        layout.addWidget(self.module_label)

        layout.addWidget(QLabel("출력 장 수 (라벨지 장 수)"))
        self.sheets_input = QLineEdit()
        layout.addWidget(self.sheets_input)

        layout.addWidget(QLabel("프린터 선택"))
        self.printer_combo = QComboBox()
        self.load_printers()
        layout.addWidget(self.printer_combo)

        self.round_label   = QLabel("")
        self.perpage_label = QLabel("")
        layout.addWidget(self.round_label)
        layout.addWidget(self.perpage_label)

        self.print_btn = QPushButton("출력")
        self.print_btn.clicked.connect(self.handle_print)
        layout.addWidget(self.print_btn)

        self.setLayout(layout)
        self.update_info()

    def load_printers(self):
        printers = win32print.EnumPrinters(2)
        for p in printers:
            self.printer_combo.addItem(p[2])

    def update_info(self):
        cache    = load_cache()
        spec     = LABEL_SPECS[self.label_combo.currentText()]
        per_page = spec["cols"] * spec["rows"]
        self.round_label.setText(f"현재 인쇄 회차: {cache['print_round']}회")
        self.perpage_label.setText(
            f"선택 라벨: {spec['cols']}열 × {spec['rows']}행 = 장당 {per_page}개"
        )

    def handle_print(self):
        spec_name = self.label_combo.currentText()
        spec      = LABEL_SPECS[spec_name]
        per_page  = spec["cols"] * spec["rows"]

        site    = self.site
        module  = self.module
        printer = self.printer_combo.currentText()

        try:
            sheets = int(self.sheets_input.text())
            if sheets <= 0:
                raise ValueError
        except Exception:
            QMessageBox.warning(self, "오류", "장 수를 올바르게 입력하세요.")
            return

        total_qty = sheets * per_page

        cache         = load_cache()
        today         = cache["date"]
        task_index    = cache["daily_task_index"]
        current_round = cache["print_round"]

        serial_key  = f"{site}_{module}_{current_round}"
        last_serial = cache["serial"].get(serial_key, -1)

        temp_task   = task_index + 1
        temp_serial = last_serial

        labels = []
        for _ in range(total_qty):
            temp_serial += 1
            labels.append(f"KR-{site}{module}{temp_task}-{today}-{temp_serial:03d}")

        confirm = QMessageBox.question(
            self,
            "최종 확인",
            f"{labels[0]}\n~\n{labels[-1]}\n\n"
            f"{sheets}장 출력 맞으십니까?\n"
            f"(QR 개수: {total_qty}개  |  {spec_name})",
            QMessageBox.Yes | QMessageBox.No
        )
        if confirm == QMessageBox.No:
            return

        try:
            print_gdi(labels, printer, spec)
        except Exception as e:
            QMessageBox.critical(self, "출력 오류", str(e))
            return

        cache["daily_task_index"]   = task_index + 1
        cache["serial"][serial_key] = temp_serial
        cache["print_round"]       += 1
        save_cache(cache)

        self.update_info()
        QMessageBox.information(self, "완료", f"{sheets}장 ({total_qty}개) 출력 완료")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = LabelApp()
    window.show()
    sys.exit(app.exec_())