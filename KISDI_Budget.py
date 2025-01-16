import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.detach(), encoding='utf-8')

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import re

def format_number(num):
    """숫자를 포맷하는 함수."""
    if isinstance(num, float) and num.is_integer():
        return str(int(num))
    elif isinstance(num, (int, float)):
        return f"{num:,.0f}"
    else:
        return ""

def make_expression(item_name, unit_price, qty, qty_unit, freq1, freq1_unit, freq2, freq2_unit, amount):
    """항목의 상세 내역을 조합합니다."""
    parts = []
    if pd.notna(unit_price):
        parts.append(str(format_number(unit_price)))
    if pd.notna(qty) and qty != 0:
        if pd.notna(qty_unit):
            parts.append(f"{int(qty)}{qty_unit}")
        else:
            parts.append(str(qty))
    if pd.notna(freq1) and freq1 != 0:
        if pd.notna(freq1_unit):
            parts.append(f"{format_number(freq1)}{freq1_unit}")
        else:
            parts.append(str(format_number(freq1)))
    if pd.notna(freq2) and freq2 != 0:
        if pd.notna(freq2_unit):
            parts.append(f"{format_number(freq2)}{freq2_unit}")
        else:
            parts.append(str(format_number(freq2)))

    expr_str = "×".join(parts)
    amount_str = int(amount)

    return f"{item_name} : {expr_str}={amount_str}"

def parse_excel(file_path):
    """
    엑셀 파일을 파싱하여 중분류와 세부항목을 추출합니다.
    """
    try:
        df = pd.read_excel(file_path, header=None)
        print("엑셀 파일을 성공적으로 읽었습니다.")
    except Exception as e:
        messagebox.showerror("파일 오류", f"엑셀 파일을 읽는 중 오류가 발생했습니다.\n{e}")
        return []

    # 필요한 열 선택 및 이름 부여
    try:
        needed = df[[1,3,4,5,6,7,8,9,10,11]].copy()
        needed.columns = [
            "raw_cat",     # B열
            "항목명",      # D열
            "단가",       # E열
            "갯수",       # F열
            "갯수단위",    # G열
            "횟수1",      # H열
            "횟수1단위",   # I열
            "횟수2",      # J열
            "횟수2단위",   # K열
            "금액"        # L열
        ]
    except KeyError:
        messagebox.showerror("열 오류", "엑셀 파일의 열 구조가 예상과 다릅니다.\n필요한 열(B, D, E, F, G, H, I, J, K, L)이 모두 있는지 확인하세요.")
        return []

    # 숫자 변환
    for col in ["단가", "갯수", "횟수1", "횟수2", "금액"]:
        needed[col] = pd.to_numeric(needed[col], errors="coerce")

    # 중분류와 항목 파싱
    cat_counter = 0
    group_dict = {}
    last_main_item = None

    for idx, row in needed.iterrows():
        raw_cat = str(row["raw_cat"]).strip() if pd.notna(row["raw_cat"]) else ""

        if re.match(r"^\d+\)", raw_cat):  # 중분류인지 확인
            cat_counter += 1
            middle_name = re.sub(r"^\d+\)", "", raw_cat).strip()
            final_label = f"{cat_counter}. {middle_name}"
            if final_label not in group_dict:
                group_dict[final_label] = {"items": [], "total": 0}
            continue

        # 현재 카테고리가 없는 경우 처리하지 않음
        if not group_dict:
            continue

        current_cat = list(group_dict.keys())[-1]

        # 항목 세부정보 가져오기
        item_name = str(row["항목명"]).strip()

        # "-"로 시작하는 항목명도 그대로 사용하도록 수정
        # if item_name.startswith("-"):
        #    if last_main_item is not None:
        #        item_name = last_main_item

        last_main_item = item_name  # 현재 항목을 마지막 주요 항목으로 설정

        unit_price = row["단가"]
        qty = row["갯수"]
        qty_unit = row["갯수단위"]
        freq1 = row["횟수1"]
        freq1_unit = row["횟수1단위"]
        freq2 = row["횟수2"]
        freq2_unit = row["횟수2단위"]
        amount = row["금액"]

        # 금액이 있는 항목만 처리
        if pd.notna(amount) and amount != 0:
            expr_str = make_expression(
                item_name,
                unit_price,
                qty,
                qty_unit,
                freq1,
                freq1_unit,
                freq2,
                freq2_unit,
                amount
            )
            group_dict[current_cat]["items"].append(expr_str)
            group_dict[current_cat]["total"] += amount

    # 최종 데이터 변환
    results = []
    for cat, info in group_dict.items():
        total = info["total"]
        total_val = int(total) if total != 0 else ""
        results.append({
            "구분": cat,
            "내용": "\n".join(info["items"]),
            "금액": total_val
        })

    return results


def build_final_report(parsed_list):
    """
    최종 보고서 텍스트를 생성합니다.
    """
    if not parsed_list:
        return "데이터가 없습니다."

    lines = []
    for item in parsed_list:
        cat = item["구분"]  # e.g., "1. 사업인건비"
        content = item["내용"]  # e.g., "- 인쇄비 : 50×10월=500\n- 복사료 : ..."
        amount = item["금액"]  # e.g., 500 or ""

        # If amount is not empty, format with commas
        amount_str = format_number(amount)

        # Combine into desired format
        line = f"{cat} | {content} | {amount_str}"
        lines.append(line)

    # Join all categories with double newline
    report_text = "\n\n".join(lines)
    return report_text

# ------------------ GUI ------------------ #
# ------------------ GUI ------------------ #
class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("예산 보고서 추출기")
        self.geometry("1000x700")
        
        # 파일 경로 표시
        lbl_file = tk.Label(self, text="엑셀 파일 경로:")
        lbl_file.pack(anchor="w", padx=10, pady=(10,0))
        
        self.var_path = tk.StringVar()
        ent_file = tk.Entry(self, textvariable=self.var_path, width=100)
        ent_file.pack(anchor="w", padx=10)
        
        # 버튼들
        frm_btns = tk.Frame(self)
        frm_btns.pack(anchor="w", padx=10, pady=5)
        
        btn_select = tk.Button(frm_btns, text="파일 선택", command=self.select_file)
        btn_select.pack(side="left")
        
        btn_clear = tk.Button(frm_btns, text="초기화", command=self.clear_all)
        btn_clear.pack(side="left", padx=5)
        
        # 추가된 버튼들
        btn_export_excel = tk.Button(frm_btns, text="엑셀로 내보내기", command=self.export_to_excel)
        btn_export_excel.pack(side="left", padx=5)
        
        btn_copy_clipboard = tk.Button(frm_btns, text="클립보드로 복사", command=self.copy_to_clipboard)
        btn_copy_clipboard.pack(side="left", padx=5)
        
        # Treeview: 3열 (중분류, 내용, 금액)
        columns = ("col_cat", "col_desc", "col_amount")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("col_cat", text="중분류")
        self.tree.heading("col_desc", text="내용")
        self.tree.heading("col_amount", text="금액")
        
        self.tree.column("col_cat", width=200, anchor="w")
        self.tree.column("col_desc", width=600, anchor="w")
        self.tree.column("col_amount", width=150, anchor="e")
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Scrollbars for Treeview
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        vsb.place(x=900, y=100, height=500)
        self.tree.configure(yscrollcommand=vsb.set)
        
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        hsb.pack(side="bottom", fill="x")
        self.tree.configure(xscrollcommand=hsb.set)
        
        # 최종 텍스트
        frame_text = tk.LabelFrame(self, text="최종 보고서 텍스트")
        frame_text.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.txt_report = tk.Text(frame_text, wrap="none")
        self.txt_report.pack(fill="both", expand=True)
        
        # Scrollbars for Text widget
        txt_vsb = ttk.Scrollbar(frame_text, orient="vertical", command=self.txt_report.yview)
        txt_vsb.pack(side="right", fill="y")
        self.txt_report.configure(yscrollcommand=txt_vsb.set)
        
        txt_hsb = ttk.Scrollbar(frame_text, orient="horizontal", command=self.txt_report.xview)
        txt_hsb.pack(side="bottom", fill="x")
        self.txt_report.configure(xscrollcommand=txt_hsb.set)
        
        # 안내 라벨
        lbl_info = tk.Label(self, text="사용법: [파일 선택] → 엑셀 지정 → 결과 확인\nPyInstaller 등으로 exe 빌드 가능.",
                            fg="blue")
        lbl_info.pack(padx=10, pady=5)
        
        self.parsed_data = []
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            messagebox.showinfo("정보", "실패")
            return
        self.var_path.set(file_path)
        self.parse_and_show(file_path)
    
    def parse_and_show(self, file_path):
        self.clear_all()
        
        # 파싱
        self.parsed_data = parse_excel(file_path)
        
        if not self.parsed_data:
            messagebox.showinfo("정보", "파싱된 데이터가 없습니다.")
            return
        
        # Treeview 표시
        for row in self.parsed_data:
            cat = row["구분"]       # e.g., "1. 사업인건비"
            desc = row["내용"]     # e.g., "- 인쇄비 : 50×10월=500\n- 복사료 : ..."
            amt = row["금액"]      # e.g., 500 or ""
            
            if desc:
                lines = desc.split('\n')
                for idx, line in enumerate(lines):
                    if idx == 0:
                        # 첫 번째 항목과 함께 중분류 표시
                        self.tree.insert("", "end", values=(cat, line, format_number(amt)))
                    else:
                        # 이후 항목은 중분류 없이 내용과 금액만 표시
                        self.tree.insert("", "end", values=("", line, ""))
            else:
                # 중분류만 표시하고 내용과 금액은 비워둠
                self.tree.insert("", "end", values=(cat, "", ""))
        
        # 최종 보고서 텍스트
        final_text = build_final_report(self.parsed_data)
        self.txt_report.insert("1.0", final_text)
    
    def clear_all(self):
        # Treeview 비우기
        for item in self.tree.get_children():
            self.tree.delete(item)
        # Text 비우기
        self.txt_report.delete("1.0", "end")
        self.parsed_data = []
    
    def export_to_excel(self):
        if not self.parsed_data:
            messagebox.showinfo("정보", "내보낼 데이터가 없습니다.")
            return
        
        # DataFrame으로 변환
        df = pd.DataFrame(self.parsed_data)
        
        # 파일 저장 대화상자
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="엑셀 파일로 내보내기"
        )
        if not file_path:
            return
        
        try:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("성공", f"데이터를 성공적으로 '{file_path}'에 저장했습니다.")
        except Exception as e:
            messagebox.showerror("저장 오류", f"엑셀 파일로 저장하는 중 오류가 발생했습니다.\n{e}")
    
    def copy_to_clipboard(self):
        if not self.parsed_data:
            messagebox.showinfo("정보", "복사할 데이터가 없습니다.")
            return
        
        # DataFrame으로 변환
        df = pd.DataFrame(self.parsed_data)
        
        # 탭으로 구분된 문자열 생성
        data_str = df.to_csv(sep='\t', index=False)
        
        try:
            pyperclip.copy(data_str)
            messagebox.showinfo("성공", "데이터가 클립보드에 복사되었습니다.")
        except Exception as e:
            messagebox.showerror("복사 오류", f"클립보드로 복사하는 중 오류가 발생했습니다.\n{e}")

if __name__ == "__main__":
    # pyperclip 설치 여부 확인
    try:
        import pyperclip
    except ImportError:
        messagebox.showerror("라이브러리 오류", "pyperclip 라이브러리가 설치되지 않았습니다.\n명령어: pip install pyperclip")
        sys.exit(1)
    
    app = MyApp()
    app.mainloop()
