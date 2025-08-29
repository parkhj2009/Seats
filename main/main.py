# 2025년도에 만든 교실 자리 뽑기 프로그램 V2
import tkinter as tk
from tkinter import *
import random as r
from tkinter import messagebox, filedialog
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.utils import range_boundaries
import os

# 전역 변수
excluded = set()  # 제외할 번호
selected = set()  # 비활성화된 자리 번호
seat_buttons = []  # 자리 버튼들
is_seat_creation_phase = False  # 자리 생성 단계인지 여부
first_selected_seat = None  # 첫 번째 선택된 자리
current_seat_assignment = {}  # 현재 자리 배정 상태

def toggle_exclude(num, button):
    if num in excluded:
        excluded.remove(num)
        button.config(bg='lightgray', text='')
    else:
        excluded.add(num)
        button.config(bg='red', text='X')

def add_excluded_numbers():
    try:
        # 기존 제외 목록 초기화
        excluded.clear()
        
        # 입력된 번호들을 처리
        numbers = entry_exclude.get().strip()
        if numbers:
            # 쉼표로 구분된 번호들을 처리
            for num in numbers.split(','):
                num = num.strip()
                if num:
                    num = int(num)
                    if num <= 0:
                        messagebox.showerror("오류", "1 이상의 숫자만 입력 가능합니다!")
                        return False
                    excluded.add(num)
        
        # 제외된 번호가 전체 학생 수보다 많으면 경고
        if len(excluded) > int(entry_students.get()):
            messagebox.showerror("오류", "제외할 번호가 전체 학생 수보다 많습니다!")
            return False
            
        return True
    except ValueError:
        messagebox.showerror("오류", "올바른 숫자를 입력해주세요!")
        return False

def generate_candidate_buttons():
    global seat_buttons, selected, is_seat_creation_phase
    for widget in frame.winfo_children():
        widget.destroy()
    seat_buttons = []
    selected = set()  # 비활성화된 자리 초기화
    is_seat_creation_phase = True  # 자리 생성 단계 시작

    try:
        nums = int(entry_students.get())
        if nums <= 0:
            messagebox.showerror("오류", "올바른 학생 수를 입력해주세요!")
            return
        if nums > 20:
            messagebox.showerror("오류", "학생 수는 20명 이하로만 입력 가능합니다!")
            return
    except ValueError:
        messagebox.showerror("오류", "올바른 학생 수를 입력해주세요!")
        return

    # 18개 자리로 고정 (기존 코드와 동일)
    total_seats = 18
    cols = 6  # 6열로 고정
    rows = 3  # 3행으로 고정

    # 모든 자리를 생성 (번호 없이)
    for i in range(rows):
        row_buttons = []
        for j in range(cols):
            idx = i * cols + j + 1
            if idx > total_seats:
                break
            
            btn = Button(frame, text='', width=8, height=3, font=('맑은 고딕', 12),
                         bg='lightblue', fg='black', command=lambda i=i, j=j: select_seat(i, j))
            btn.grid(row=i, column=j, padx=5, pady=5)
            row_buttons.append(btn)
        seat_buttons.append(row_buttons)

def select_seat(i, j):
    global selected, first_selected_seat
    idx = i * len(seat_buttons[0]) + j + 1
    
    # 자리 생성 단계에서는 자리 비활성화
    if is_seat_creation_phase:
        if idx in selected:
            selected.remove(idx)
            seat_buttons[i][j].config(bg='lightblue', text='')
        else:
            selected.add(idx)
            seat_buttons[i][j].config(bg='lightgray', text='X', fg='black')
    # 자리 배치 단계에서는 자리 교환
    else:
        # 비활성화된 자리(X)는 선택할 수 없음
        if seat_buttons[i][j]['text'] == 'X':
            return
            
        if first_selected_seat is None:
            first_selected_seat = (i, j)
            seat_buttons[i][j].config(bg='yellow')
        else:
            # 두 번째 자리 선택 시 교환
            i1, j1 = first_selected_seat
            # 첫 번째 선택된 자리의 텍스트와 배경색 저장
            temp_text = seat_buttons[i1][j1]['text']
            temp_bg = seat_buttons[i1][j1]['bg']
            
            # 두 자리의 텍스트와 배경색 교환
            seat_buttons[i1][j1].config(text=seat_buttons[i][j]['text'], bg='lightblue')
            seat_buttons[i][j].config(text=temp_text, bg='lightblue')
            
            # 자리 배정 상태 업데이트
            if temp_text and seat_buttons[i][j]['text']:
                current_seat_assignment[temp_text] = (i, j)
                current_seat_assignment[seat_buttons[i][j]['text']] = (i1, j1)
            
            # 첫 번째 선택 초기화
            first_selected_seat = None

def generate_seats():
    global seat_buttons, selected, is_seat_creation_phase, first_selected_seat, current_seat_assignment
    for widget in frame.winfo_children():
        widget.destroy()
    seat_buttons = []
    is_seat_creation_phase = False  # 자리 생성 단계 종료
    first_selected_seat = None  # 첫 번째 선택된 자리 초기화
    current_seat_assignment.clear()  # 자리 배정 상태 초기화

    try:
        nums = int(entry_students.get())
        if nums <= 0:
            messagebox.showerror("오류", "올바른 학생 수를 입력해주세요!")
            return
    except ValueError:
        messagebox.showerror("오류", "올바른 학생 수를 입력해주세요!")
        return

    # 제외할 번호 추가 및 검증
    if not add_excluded_numbers():
        return

    # 18개 자리로 고정 (기존 코드와 동일)
    total_seats = 18
    cols = 6  # 6열로 고정
    rows = 3  # 3행으로 고정

    # 제외된 학생을 제외한 학생 리스트 생성
    available_students = [i for i in range(1, nums + 1) if i not in excluded]
    r.shuffle(available_students)

    # 활성화된 자리 수와 배정할 학생 수가 일치하는지 확인
    active_seats = total_seats - len(selected)
    if active_seats != len(available_students):
        messagebox.showerror("오류", "활성화된 자리 수와 배정할 학생 수가 일치하지 않습니다!")
        return

    # 모든 자리를 생성하고 학생 배정
    student_idx = 0
    for i in range(rows):
        row_buttons = []
        for j in range(cols):
            idx = i * cols + j + 1
            if idx > total_seats:
                break
            
            # 현재 자리가 비활성화된 자리인 경우
            if idx in selected:
                btn = Button(frame, text='X', width=8, height=3, font=('맑은 고딕', 12),
                             bg='lightgray', fg='black', state='disabled')
            # 현재 자리에 배정할 학생이 있는 경우
            elif student_idx < len(available_students):
                student = available_students[student_idx]
                student_idx += 1
                btn = Button(frame, text=str(student), width=8, height=3, font=('맑은 고딕', 12),
                             bg='lightblue', fg='black', command=lambda i=i, j=j: select_seat(i, j))
                current_seat_assignment[str(student)] = (i, j)
            else:
                btn = Button(frame, text='', width=8, height=3, font=('맑은 고딕', 12),
                             bg='lightgray', fg='black', state='disabled')

            btn.grid(row=i, column=j, padx=5, pady=5)
            row_buttons.append(btn)
        seat_buttons.append(row_buttons)

def create_excel_file():
    # 엑셀 파일 생성 함수
    try:
        # 입력값 검증
        grade = int(entry_grade.get())
        group = int(entry_group.get())
        n = int(entry_students.get())
        teacher = entry_teacher.get().strip()
        
        if not teacher:
            messagebox.showerror("오류", "담임선생님 성함을 입력해주세요!")
            return
            
        if not current_seat_assignment:
            messagebox.showerror("오류", "먼저 자리 배치를 완료해주세요!")
            return
            
    except ValueError:
        messagebox.showerror("오류", "올바른 숫자를 입력해주세요!")
        return

    # 바탕화면 경로 설정
    desktop_path = os.path.expanduser("~/Desktop")
    default_filename = f"{grade}학년{group}반_좌석배정표.xlsx"
    default_path = os.path.join(desktop_path, default_filename)

    # 파일 저장 위치 선택
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="엑셀 파일 저장",
        initialdir=desktop_path,
        initialfile=default_filename
    )
    
    if not file_path:
        return

    # 엑셀 파일 생성
    xlsx = Workbook()
    x1 = xlsx.active

    # === 인쇄 설정 추가 ===
    x1.page_setup.paperSize = x1.PAPERSIZE_A4
    x1.page_setup.orientation = 'landscape'
    x1.page_margins.left = 1.0
    x1.page_margins.right = 1.0
    x1.page_margins.top = 1.0
    x1.page_margins.bottom = 1.0
    x1.page_margins.header = 0.5
    x1.page_margins.footer = 0.5
    # =====================

    # 폰트 스타일 정의
    Title_font = Font(name='Pretendard', size=24, bold=True)
    Pretendard = Font(name='Pretendard', size=12, bold=True)

    # 테두리 스타일 정의
    Thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    Thick_border = Side(style='thick')
    No_border = Side(style=None)

    # 전체 범위 지정
    min_row, max_row = 1, 26  # 행
    min_col, max_col = 1, 13  # 열

    # 각 셀에 대해 위치에 따라 테두리 적용
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = x1.cell(row=row, column=col)

            # 각 방향 테두리 조건부 설정
            if row == min_row:
                top = Thick_border
                bottom = No_border
            elif row == max_row:
                top = No_border
                bottom = Thick_border
            else:
                top = No_border
                bottom = No_border

            if col == min_col:
                left = Thick_border
                right = No_border
            elif col == max_col:
                left = No_border
                right = Thick_border
            else:
                left = No_border
                right = No_border

            cell.border = Border(top=top, bottom=bottom, left=left, right=right)
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # 기본 병합
    x1.merge_cells('B2:L3')  # 좌석 배정표
    x1.merge_cells('B22:C22')  # 학반
    x1.merge_cells('B23:C24')  # 담임선생님
    x1.merge_cells('E23:I24')  # 칠판

    # 기본 테두리
    for row in x1['B2:L3']:  # 좌석 배치표(타이틀)
        for cell in row:
            cell.border = Thin_border

    for row in x1['B22:C22']:  # 학반
        for cell in row:
            cell.border = Thin_border

    for row in x1['B23:C24']:  # 담임 선생님
        for cell in row:
            cell.border = Thin_border

    for row in x1['E23:I24']:  # 칠판
        for cell in row:
            cell.border = Thin_border

    # 기본 데이터 입력
    x1['B2'] = "좌석 배정표"
    x1['B2'].font = Title_font

    x1['E23'] = "칠판"
    x1['E23'].font = Pretendard

    # 자리표 배치 좌표 (기존 코드와 동일)
    seat_positions = [
        ('K17:L19', 17, 11),  # 1
        ('H17:I19', 17, 8),   # 2
        ('E17:F19', 17, 5),   # 3
        ('B17:C19', 17, 2),   # 4
        ('K14:L16', 14, 11),  # 5
        ('H14:I16', 14, 8),   # 6
        ('E14:F16', 14, 5),   # 7
        ('B14:C16', 14, 2),   # 8
        ('K11:L13', 11, 11),  # 9
        ('H11:I13', 11, 8),   # 10
        ('E11:F13', 11, 5),   # 11
        ('B11:C13', 11, 2),   # 12
        ('K8:L10', 8, 11),    # 13
        ('H8:I10', 8, 8),     # 14
        ('E8:F10', 8, 5),     # 15
        ('B8:C10', 8, 2),     # 16
        ('K5:L7', 5, 11),     # 17
        ('H5:I7', 5, 8),      # 18
        ('E5:F7', 5, 5),      # 19
        ('B5:C7', 5, 2)       # 20
    ]

    # 학생 리스트 준비 (제외 번호 제외)
    # students = [str(i) for i in range(1, n + 1) if i not in excluded]
    # student_idx = 0
    for idx, (merge_range, row, col) in enumerate(seat_positions):
        gui_row = idx // 4
        gui_col = idx % 4
        if seat_buttons and seat_buttons[gui_row][gui_col]['text'] == 'X':
            continue
        button_text = seat_buttons[gui_row][gui_col]['text']
        if button_text and button_text != 'X':
            x1.merge_cells(merge_range)
            x1.cell(row=row, column=col).value = button_text
            set_border_to_merged_range(x1, merge_range, Thin_border)
        # 학생이 없으면 빈 칸(아무것도 안함)

    # 추가 데이터 입력
    x1['B22'] = f'{grade}-{group}'
    x1['B22'].font = Pretendard

    x1['B23'] = teacher
    x1['B23'].font = Pretendard

    # 엑셀 파일 저장
    try:
        xlsx.save(file_path)
        messagebox.showinfo("성공", f"엑셀 파일이 성공적으로 저장되었습니다!\n저장 위치: {file_path}")
    except Exception as e:
        messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다: {str(e)}")

def can_assign_seats():
    try:
        nums = int(entry_students.get())
        if nums <= 0:
            messagebox.showerror("오류", "올바른 학생 수를 입력해주세요!")
            generate_candidate_buttons()
            return False
    except ValueError:
        messagebox.showerror("오류", "올바른 학생 수를 입력해주세요!")
        generate_candidate_buttons()
        return False

    if not add_excluded_numbers():
        generate_candidate_buttons()
        return False

    total_seats = 18
    active_seats = total_seats - len(selected)
    available_students = [i for i in range(1, nums + 1) if i not in excluded]
    if active_seats != len(available_students):
        messagebox.showerror("오류", "활성화된 자리 수와 배정할 학생 수가 일치하지 않습니다!")
        generate_candidate_buttons()
        return False
    return True

def start_countdown_and_generate_seats():
    if not can_assign_seats():
        return
    set_inputs_state('disabled')
    countdown_label.config(text='3')
    root.after(700, lambda: countdown_label.config(text='2'))
    root.after(1400, lambda: countdown_label.config(text='1'))
    root.after(2100, lambda: [countdown_label.config(text=''), generate_seats(), set_inputs_state('normal')])

def set_border_to_merged_range(ws, merge_range, border):
    min_col, min_row, max_col, max_row = range_boundaries(merge_range)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = border

# 메인 윈도우 생성
root = Tk()
root.title("교실 자리 배치 프로그램")
root.geometry("1000x800")
root.config(bg='white')

# 입력 프레임 생성
input_frame = Frame(root, bg='white')
input_frame.grid(row=0, column=0, columnspan=3, padx=20, pady=20, sticky='ew')

# 입력 필드들
label_grade = Label(input_frame, text='학년', bg='white', fg='black', font=('맑은 고딕', 12, 'bold'))
label_grade.grid(row=0, column=0, padx=10, pady=5, sticky='e')
entry_grade = Entry(input_frame, width=15, font=('맑은 고딕', 12), bd=1, relief='solid', bg='white', fg='black')
entry_grade.grid(row=0, column=1, padx=10, pady=5)

label_group = Label(input_frame, text='반', bg='white', fg='black', font=('맑은 고딕', 12, 'bold'))
label_group.grid(row=0, column=2, padx=10, pady=5, sticky='e')
entry_group = Entry(input_frame, width=15, font=('맑은 고딕', 12), bd=1, relief='solid', bg='white', fg='black')
entry_group.grid(row=0, column=3, padx=10, pady=5)

label_students = Label(input_frame, text='학생 수\n(1~18)', bg='white', fg='black', font=('맑은 고딕', 12, 'bold'))
label_students.grid(row=1, column=0, padx=10, pady=5, sticky='e')
entry_students = Entry(input_frame, width=15, font=('맑은 고딕', 12), bd=1, relief='solid', bg='white', fg='black')
entry_students.grid(row=1, column=1, padx=10, pady=5)

label_teacher = Label(input_frame, text='담임선생님', bg='white', fg='black', font=('맑은 고딕', 12, 'bold'))
label_teacher.grid(row=1, column=2, padx=10, pady=5, sticky='e')
entry_teacher = Entry(input_frame, width=15, font=('맑은 고딕', 12), bd=1, relief='solid', bg='white', fg='black')
entry_teacher.grid(row=1, column=3, padx=10, pady=5)

label_exclude = Label(input_frame, text='제외할 번호\n(쉼표로 구분)', bg='white', fg='black', font=('맑은 고딕', 12, 'bold'))
label_exclude.grid(row=2, column=0, padx=10, pady=5, sticky='e')
entry_exclude = Entry(input_frame, width=15, font=('맑은 고딕', 12), bd=1, relief='solid', bg='white', fg='black')
entry_exclude.grid(row=2, column=1, padx=10, pady=5)

# 버튼들
btn_frame = Frame(input_frame, bg='white')
btn_frame.grid(row=2, column=2, columnspan=2, padx=10, pady=5)

btn_generate_candidates = Button(btn_frame, text='자리 생성', 
                               command=generate_candidate_buttons,
                               font=('맑은 고딕', 11, 'bold'), bg='#4CAF50', fg='#000000',
                               relief='raised', bd=2, width=10)
btn_generate_candidates.grid(row=0, column=0, padx=5, pady=5)

btn_generate_seats = Button(btn_frame, text='자리 배치', 
                          command=start_countdown_and_generate_seats,
                          font=('맑은 고딕', 11, 'bold'), bg='#2196F3', fg='#000000',
                          relief='raised', bd=2, width=10)
btn_generate_seats.grid(row=0, column=1, padx=5, pady=5)

btn_create_excel = Button(btn_frame, text='엑셀 생성', 
                         command=create_excel_file,
                         font=('맑은 고딕', 11, 'bold'), bg='#FF9800', fg='#000000',
                         relief='raised', bd=2, width=10)
btn_create_excel.grid(row=0, column=2, padx=5, pady=5)

# 입력 필드와 버튼을 리스트로 관리
all_inputs = [
    entry_grade, entry_group, entry_students, entry_teacher, entry_exclude,
    btn_generate_candidates, btn_generate_seats, btn_create_excel
]

def set_inputs_state(state):
    for widget in all_inputs:
        widget.config(state=state)

# 설명 라벨
info_label = Label(input_frame, text="사용법: 1. 정보 입력 → 2. 자리 생성 → 3. 비활성화할 자리 선택 → 4. 자리 배치 → 5. 엑셀 생성", 
                  bg='white', fg='#666666', font=('맑은 고딕', 20))
info_label.grid(row=3, column=0, columnspan=4, pady=10)

# 칠판 위치 표시 라벨
blackboard_label = Button(input_frame, text="칠판", 
                        font=('맑은 고딕', 11, 'bold'), bg='#FF9800', fg='#000000',
                         relief='raised', bd=2, width=100)
blackboard_label.grid(row=4, column=0, columnspan=4, pady=5)

# 자리 배치 프레임
frame = Frame(root, bg='white')
frame.grid(row=1, column=0, columnspan=3, padx=20, pady=20)

# 카운트다운 라벨 추가
countdown_label = Label(root, text='', font=('맑은 고딕', 40, 'bold'), bg='white', fg='red')
countdown_label.grid(row=2, column=0, columnspan=3, pady=10)

root.mainloop()