import tkinter as tk
from tkinter import *
import random as r
from tkinter import messagebox

# 전역 변수
excluded = set()  # 제외할 번호
selected = set()  # 비활성화된 자리 번호
seat_buttons = []  # 자리 버튼들
is_seat_creation_phase = False  # 자리 생성 단계인지 여부
first_selected_seat = None  # 첫 번째 선택된 자리
last_input_values = None  # 마지막 입력값 저장

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
        numbers = entry3.get().strip()
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
        if len(excluded) > int(entry1.get()):
            messagebox.showerror("오류", "제외할 번호가 전체 학생 수보다 많습니다!")
            return False
            
        return True
    except ValueError:
        messagebox.showerror("오류", "올바른 숫자를 입력해주세요!")
        return False

def generate_candidate_buttons():
    global seat_buttons, selected, is_seat_creation_phase, first_selected_seat, last_input_values
    for widget in frame.winfo_children():
        widget.destroy()
    seat_buttons = []
    selected = set()  # 비활성화된 자리 초기화
    is_seat_creation_phase = True  # 자리 생성 단계 시작
    first_selected_seat = None  # 첫 번째 선택된 자리 초기화
    
    try:
        nums = int(entry1.get())
        cols = int(entry2.get())
        if nums <= 0 or cols <= 0:
            messagebox.showerror("오류", "올바른 숫자를 입력해주세요!")
            return
    except ValueError:
        messagebox.showerror("오류", "올바른 숫자를 입력해주세요!")
        return

    # 현재 입력값 저장
    last_input_values = (nums, cols, entry3.get().strip())

    # 전체 자리 수 계산
    total_seats = nums
    rows = (total_seats + cols - 1) // cols

    # 모든 자리를 생성 (번호 없이)
    for i in range(rows):
        row_buttons = []
        for j in range(cols):
            idx = i * cols + j + 1
            if idx > total_seats:
                break
            
            btn = Button(frame, text='', width=5, height=2, font=('맑은 고딕', 14),
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
        print(f"Selected seats: {selected}")  # 디버그 메시지
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
            
            # 두 자리의 텍스트 교환
            seat_buttons[i1][j1].config(text=seat_buttons[i][j]['text'], bg='lightblue')
            seat_buttons[i][j].config(text=temp_text, bg='lightblue')
            
            # 첫 번째 선택 초기화
            first_selected_seat = None

def generate_seats():
    global seat_buttons, selected, first_selected_seat
    
    # 자리 버튼이 없는 경우에만 에러 메시지 출력
    if not seat_buttons:
        messagebox.showerror("오류", "먼저 '자리 생성' 버튼을 눌러주세요!")
        return
        
    for widget in frame.winfo_children():
        widget.destroy()
    seat_buttons = []
    first_selected_seat = None  # 첫 번째 선택된 자리 초기화

    try:
        nums = int(entry1.get())
        cols = int(entry2.get())
        if nums <= 0 or cols <= 0:
            messagebox.showerror("오류", "올바른 숫자를 입력해주세요!")
            return
    except ValueError:
        messagebox.showerror("오류", "올바른 숫자를 입력해주세요!")
        return

    # 제외할 번호 추가 및 검증
    if not add_excluded_numbers():
        return

    # 전체 자리 수 계산
    total_seats = nums
    rows = (total_seats + cols - 1) // cols

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
                btn = Button(frame, text='X', width=5, height=2, font=('맑은 고딕', 14),
                             bg='lightgray', fg='black', state='disabled')
            # 현재 자리에 배정할 학생이 있는 경우
            elif student_idx < len(available_students):
                student = available_students[student_idx]
                student_idx += 1
                btn = Button(frame, text=str(student), width=5, height=2, font=('맑은 고딕', 14),
                             bg='lightblue', fg='black', command=lambda i=i, j=j: select_seat(i, j))
            else:
                btn = Button(frame, text='', width=5, height=2, font=('맑은 고딕', 14),
                             bg='lightgray', fg='black', state='disabled')

            btn.grid(row=i, column=j, padx=5, pady=5)
            row_buttons.append(btn)
        seat_buttons.append(row_buttons)

# 메인 윈도우 생성
tk = Tk()
tk.title("교실 자리 배치")
tk.geometry("800x600")
tk.config(bg='white')

# 입력 프레임 생성
input_frame = Frame(tk, bg='white')
input_frame.grid(row=0, column=0, columnspan=3, padx=20, pady=20, sticky='ew')

# 입력 필드들
label1 = Label(input_frame, text='학생 수', bg='white', fg='black', font=('맑은 고딕', 14, 'bold'))
label1.grid(row=0, column=0, padx=10, pady=10, sticky='e')
entry1 = Entry(input_frame, width=20, font=('맑은 고딕', 14), bd=1, relief='solid', bg='white', fg='black')
entry1.grid(row=0, column=1, padx=10, pady=10)

label2 = Label(input_frame, text='열 수', bg='white', fg='black', font=('맑은 고딕', 14, 'bold'))
label2.grid(row=1, column=0, padx=10, pady=10, sticky='e')
entry2 = Entry(input_frame, width=20, font=('맑은 고딕', 14), bd=1, relief='solid', bg='white', fg='black')
entry2.grid(row=1, column=1, padx=10, pady=10)

label3 = Label(input_frame, text='제외할 번호\n(쉼표로 구분)', bg='white', fg='black', font=('맑은 고딕', 14, 'bold'))
label3.grid(row=2, column=0, padx=10, pady=10, sticky='e')
entry3 = Entry(input_frame, width=20, font=('맑은 고딕', 14), bd=1, relief='solid', bg='white', fg='black')
entry3.grid(row=2, column=1, padx=10, pady=10)

# 버튼들
btn_generate_candidates = Button(input_frame, text='자리 생성', 
                               command=generate_candidate_buttons,
                               font=('맑은 고딕', 12, 'bold'), bg='#ffffff', fg='black',
                               relief='raised', bd=2)
btn_generate_candidates.grid(row=0, column=2, padx=10, pady=10)

btn_generate_seats = Button(input_frame, text='자리 배치', 
                          command=generate_seats,
                          font=('맑은 고딕', 12, 'bold'), bg='#ffffff', fg='black',
                          relief='raised', bd=2)
btn_generate_seats.grid(row=1, column=2, padx=10, pady=10)


# 자리 배치 프레임
frame = Frame(tk, bg='white')
frame.grid(row=1, column=0, columnspan=3, padx=20, pady=20)

tk.mainloop()
