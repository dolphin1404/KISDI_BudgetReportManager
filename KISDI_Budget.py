import sys
import os

# 깃허브 주소입니다.
# https://github.com/dolphin1404/KISDI_BudgetReportManager 

# 터미널에서 디버그 한글 깨짐 현상 해결용용
#import io
#sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding='utf-8')
#sys.stderr = io.TextIOWrapper(sys.stderr.detach(), encoding='utf-8')


import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import re



def add_commas(num: float) -> str:
    """정수나 실수를 천 단위마다 콤마로 구분하여 문자열로 반환합니다."""
    # 실수이면서 소수점 이하가 0이라면 정수로 변환
    if isinstance(num, float) and num.is_integer():
        return f"{int(num):,}"
    # 정수 혹은 일반 실수라면 그대로 콤마를 붙임
    elif isinstance(num, (int, float)):
        return f"{num:,}"
    else:
        return ""

def make_expression(item_name, unit_price, qty, qty_unit, freq1, freq1_unit, freq2, freq2_unit, amount):
    """항목의 상세 내역을 조합합니다."""
    parts = []
    if pd.notna(unit_price):
        parts.append(str(add_commas(unit_price)))
    if pd.notna(qty) and qty != 0:
        if pd.notna(qty_unit):
            parts.append(f"{int(qty)}{qty_unit}")
        else:
            parts.append(str(qty))
    if pd.notna(freq1) and freq1 != 0:
        if pd.notna(freq1_unit):
            parts.append(f"{add_commas(freq1)}{freq1_unit}")
        else:
            parts.append(str(add_commas(freq1)))
    if pd.notna(freq2) and freq2 != 0:
        if pd.notna(freq2_unit):
            parts.append(f"{add_commas(freq2)}{freq2_unit}")
        else:
            parts.append(str(add_commas(freq2)))

    expr_str = "×".join(parts)
    amount_str = format(int(amount), ",d")

    return f"- {item_name} : {expr_str}={amount_str}"

def parse_excel(file_path):
    """
    엑셀 파일을 파싱하여 중분류와 세부항목을 추출
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

        if re.match(r"^\d+\)", raw_cat):  # 중분류인지 확인  ")" 닫는 괄호로 판단함
            cat_counter += 1
            middle_name = re.sub(r"^\d+\)", "", raw_cat).strip()
            final_label = f"{cat_counter}. {middle_name}" # 
            if final_label not in group_dict:
                group_dict[final_label] = {"items": [], "total": 0}
            continue

        # 현재 카테고리가 없는 경우 처리하지 않음
        if not group_dict:
            continue

        current_cat = list(group_dict.keys())[-1]

        # 항목 세부정보 가져오기
        item_name = str(row["항목명"]).strip()

        if item_name.startswith("-"):
            # 이전 항목명 상속
            if last_main_item is not None:
                item_name = last_main_item

        else:
            # 새로운 주요 항목으로 설정
            last_main_item = item_name

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
        total_val = add_commas(int(total)) if total != 0 else ""
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
        amount_str = amount

        # Combine into desired format
        line = f"{cat} | {content} | {amount_str}"
        lines.append(line)

    # Join all categories with double newline
    report_text = "\n\n".join(lines)
    return report_text

# ------------------ GUI ------------------ #
class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("예산 보고서 추출기 [충북대 소프트웨어학부 이규민, 윤준식]")
        self.geometry("800x700")
        
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
        btn_clear.pack(side ="left", padx=5)

        btn_export_excel = tk.Button(frm_btns, text="엑셀로 내보내기", command=self.export_to_excel)
        btn_export_excel.pack(side="left", padx=5)
        
        btn_export_sheet = tk.Button(frm_btns, text="기존 엑셀에 시트 추가", command=self.export_to_existing_excel)
        btn_export_sheet.pack(side="left", padx=5)

        # 안내 라벨
        lbl_info = tk.Label(self, text="<단독 파일 처리>", fg="blue", font=("굴림", 10, "bold"))
        lbl_info.pack(padx=10, pady=5)
        lbl_info = tk.Label(self, text="[파일 선택] → 엑셀 파일 선택 → 결과 확인 → [엑셀로 내보내기] 혹은 [기존 엑셀에 시트 추가] → 예산보고서 생성됨",
                            fg="black")
        lbl_info.pack(padx=10, pady=5)


        # 엑셀 추출 버튼들
        export_btns = tk.Frame(self)
        export_btns.pack(anchor="w", padx=10, pady=5)
        btn_export_sheet_multi = tk.Button(export_btns, text="여러 엑셀에 각 시트 추가", command=self.process_multiple_files)
        btn_export_sheet_multi.pack(side="left", padx=5)

        btn_export_sheet_multi_2_one = tk.Button(export_btns, text="여러 엑셀에서 한 파일로 내보내기", command=self.process_multiple_files_2_one)
        btn_export_sheet_multi_2_one.pack(side="left", padx=5)
        
        lbl_info2 = tk.Label(self, text="<복수 파일 처리>", fg="blue", font=("굴림", 10, "bold"))
        lbl_info2.pack(padx=10, pady=5)
        lbl_info2 = tk.Label(self, text="[여러 엑셀에 각 시트 추가]는 엑셀 파일 각각에 \"예산보고서\"시트를 추가합니다.\n\n[여러 엑셀에서 한 파일로 내보내기]는 먼저 저장할 파일 이름을 정한 후, \n여러 엑셀 파일을 선택하고 한 파일에 모든 예산보고서를 추가합니다.",
                            fg="black")
        lbl_info2.pack(padx=10, pady=5)

        # 1) Treeview와 스크롤바를 담을 프레임 생성
        frame_tree = ttk.Frame(self)
        frame_tree.pack(fill="both", expand=True, padx=10, pady=5)

        # 2) 프레임 내부를 grid로 배치할 수 있도록 설정
        frame_tree.rowconfigure(0, weight=1)      # row=0(트리뷰) 늘어나도록
        frame_tree.columnconfigure(0, weight=1)   # col=0(트리뷰) 늘어나도록

        # 3) Treeview 생성
        columns = ("col_cat", "col_desc", "col_amount")
        self.tree = ttk.Treeview(
            frame_tree, 
            columns=columns, 
            show="headings", 
            selectmode="extended"
        )
        self.tree.heading("col_cat", text="중분류")
        self.tree.heading("col_desc", text="내용")
        self.tree.heading("col_amount", text="금액")
        
        self.tree.column("col_cat", width=130, anchor="w")
        self.tree.column("col_desc", width=350, anchor="w")
        self.tree.column("col_amount", width=100, anchor="e")

        # 4) grid에 배치, 가득 채우도록 sticky 설정
        self.tree.grid(row=0, column=0, sticky="nsew")

        # 5) 스크롤바 생성 및 grid 배치
        vsb = ttk.Scrollbar(frame_tree, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns") 
        self.tree.configure(yscrollcommand=vsb.set)

        hsb = ttk.Scrollbar(frame_tree, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=hsb.set)
        
        # 최종 보고서 텍스트

        frame_text = tk.LabelFrame(self, text="최종 보고서 텍스트")
        frame_text.pack(fill="both", expand=True, padx=10, pady=5)

        # grid로 배치할 준비
        frame_text.rowconfigure(0, weight=1)     # Text가 수직으로 늘어나도록
        frame_text.columnconfigure(0, weight=1)  # Text가 수평으로 늘어나도록

        # Text 위젯 생성
        self.txt_report = tk.Text(frame_text, wrap="none")
        self.txt_report.grid(row=0, column=0, sticky="nsew")

        # 수직 스크롤바
        txt_vsb = ttk.Scrollbar(frame_text, orient="vertical", command=self.txt_report.yview)
        txt_vsb.grid(row=0, column=1, sticky="ns")
        self.txt_report.configure(yscrollcommand=txt_vsb.set)

        # 수평 스크롤바
        txt_hsb = ttk.Scrollbar(frame_text, orient="horizontal", command=self.txt_report.xview)
        txt_hsb.grid(row=1, column=0, sticky="ew")
        self.txt_report.configure(xscrollcommand=txt_hsb.set)
        
        
        
        self.parsed_data = []
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            messagebox.showinfo("정보", "실패")
            return
        self.file_path = file_path
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
                        self.tree.insert("", "end", values=(cat, line, amt))
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
        
        default_filename = "예산보고서.xlsx"  # 기본값(파일이 없는 경우 대비)
        if self.file_path:  # 파일 경로가 존재한다면
            original_filename = os.path.basename(self.file_path)  # 예: '테스트입니다.xlsx'
            filename_without_ext, ext = os.path.splitext(original_filename)  # ('테스트입니다', '.xlsx')
            default_filename = f"예산보고서_{filename_without_ext}{ext}"  # 예: '예산보고서_테스트입니다.xlsx'

        # 파일 저장 대화상자
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_filename,  # 여기서 기본 파일명 설정
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

    def export_to_existing_excel(self):
        """
        기존 엑셀 파일에 '새 시트'를 만들어 데이터를 추가로 내보내기
        """
        if not self.parsed_data:
            messagebox.showinfo("정보", "내보낼 데이터가 없습니다.")
            return
        
        if not hasattr(self, 'file_path') or not self.file_path:
            messagebox.showinfo("정보", "먼저 엑셀 파일을 선택하세요.")
            return

        # DataFrame 변환
        df = pd.DataFrame(self.parsed_data)

        sheet_name = "예산보고서" # 시트 이름은 원하시는대로 지정해주시면 될 것 같습니다.

        # 1) 기존 엑셀 파일 경로 선택
        existing_file = self.file_path  # 사용자가 이미 선택한 파일 경로
        try:
            with pd.ExcelWriter(existing_file, 
                                engine="openpyxl", 
                                mode="a", 
                                if_sheet_exists="new") as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            messagebox.showinfo("성공", f"기존 파일 '{existing_file}'에 시트 '{sheet_name}'로 저장했습니다.")
        except Exception as e:
            messagebox.showerror("저장 오류", f"엑셀 파일에 새 시트를 추가하는 중 오류가 발생했습니다.\n{e}")
    
    # 여러 파일을 선택하고 각 파일에 대해 시트를 추가하는 함수
    def process_multiple_files(self):
        # 파일 다이얼로그를 통해 여러 파일 선택
        file_paths = filedialog.askopenfilenames(title="작업할 파일들 선택", filetypes=[("Excel Files", "*.xlsx;*.xls")])
    
        for file_path in file_paths:
            self.file_path = file_path
            self.var_path.set(file_path)
            self.parse_and_show(file_path)

            #parsed_data = parse_excel(file_path)
            #self.parsed_data = parsed_data  # 각 파일의 데이터를 self.parsed_data에 저장
            self.export_to_existing_excel()  # 저장 함수 호출
    
        # 여러 엑셀 파일을 하나의 통합 파일로 저장하는 함수
    def process_multiple_files_2_one(self):
        """
        여러 파일을 선택하여 하나의 통합된 엑셀 파일에 추가
        각 파일의 이름을 각 데이터 사이에 삽입
        """
        default_filename = "통합예산보고서.xlsx"  # 기본값(파일이 없는 경우 대비)

        # 통합 엑셀 파일 경로 설정 (새로 생성)
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_filename,  # 여기서 기본 파일명 설정
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="엑셀 파일로 내보내기"
        )
        if not output_file:
            messagebox.showinfo("정보", "저장할 파일을 선택하세요.")
            return
        
        # 여러 파일을 선택
        file_paths = filedialog.askopenfilenames(title="파일 선택", filetypes=[("Excel Files", "*.xlsx;*.xls")])

        # 새로운 엑셀 파일에 데이터를 추가
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            start_row = 0  # 첫 번째 파일부터 시작
            for idx, file_path in enumerate(file_paths):
                # 각 엑셀 파일 파싱
                parsed_data = parse_excel(file_path)

                # 파일 이름을 삽입할 행 (파일 이름을 한 행으로 추가)
                file_name = os.path.basename(file_path)
                
                # 파일 이름을 DataFrame으로 변환 (한 행 데이터로 추가)
                file_name_row = pd.DataFrame([['파일 이름:', file_name]], columns=["구분", "내용"])

                # DataFrame 처리
                df = pd.DataFrame(parsed_data)

                # 엑셀에 삽입
                file_name_row.to_excel(writer, index=False, header=False, startrow=start_row, sheet_name="통합데이터")
                df.to_excel(writer, index=False, header=(start_row == 0), startrow=start_row + 1, sheet_name="통합데이터")
                
                # 파일을 저장한 후, 다음 데이터를 위한 행 이동
                start_row += len(df) + 2  # 파일 이름을 위한 1행과 데이터 행을 포함한 크기만큼 증가

        messagebox.showinfo("성공", "여러 엑셀 파일이 통합되었습니다.")


if __name__ == "__main__":
    # pyperclip 설치 여부 확인
    # 오류 발생 시, 클립보드 복사 기능이기에 사용안하시면 삭제해도 괜찮습니다.
    try:
        import pyperclip
    except ImportError:
        messagebox.showerror("라이브러리 오류", "pyperclip 라이브러리가 설치되지 않았습니다.\n명령어: pip install pyperclip")
        sys.exit(1) # 프로그램 종료 코드로 오류발생시 주석 처리 후 시도해보세요.
    
    app = MyApp()
    app.mainloop()