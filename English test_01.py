import pandas as pd
import random
import tkinter as tk
from tkinter import messagebox

# 讀取Excel資料
data = pd.read_excel("Vocabulary List.xlsx")
words = data["EnglishWord"].tolist()
sentences = data["Sentence"].tolist()
translations = data["ChineseTranslation"].tolist()

# 建立主介面
root = tk.Tk()
root.title("英文測試軟體")
root.geometry("800x600")  # 設置視窗大小為800x600像素

# 初始變數
score = 0
question_num = 0
selected_mode = None
current_answer = ""

# 顯示題目和選項
def show_question():
    global current_answer
    question_label.config(text=f"第 {question_num + 1} 題: {current_question}")
    
    # 清空舊選項按鈕
    for widget in options_frame.winfo_children():
        widget.destroy()
    
    # 根據模式顯示選項或輸入框
    if selected_mode in [1, 2]:  # 選擇題模式
        for i, choice in enumerate(current_choices):
            btn = tk.Button(options_frame, text=choice, width=30, height=2, font=("Arial", 14),
                            command=lambda choice=choice: check_answer(choice))
            btn.pack(pady=5)
    elif selected_mode == 3:  # 克漏字模式
        input_entry = tk.Entry(options_frame, width=30, font=("Arial", 14))
        input_entry.pack(pady=5)
        
        submit_btn = tk.Button(options_frame, text="提交答案", width=30, height=2, font=("Arial", 14),
                               command=lambda: check_answer(input_entry.get()))
        submit_btn.pack(pady=5)

# 檢查答案
def check_answer(selected_choice):
    global score, question_num

    if selected_choice == current_answer:
        score += 1
        messagebox.showinfo("結果", "正確!")
    else:
        messagebox.showinfo("結果", f"錯誤! 正確答案是: {current_answer}")
    
    question_num += 1
    if question_num < 20:
        next_question()
    else:
        finish_quiz()

# 結束測試並顯示成績
def finish_quiz():
    global question_num, score
    messagebox.showinfo("測試結束", f"測試結束! 你的分數是: {score} / 20 = {score / 20 * 100:.2f}%")
    question_num = 0
    score = 0
    main_menu()

# 選擇題型並開始測試
def select_mode(mode):
    global selected_mode
    selected_mode = mode
    start_quiz()

# 開始測試
def start_quiz():
    global question_num, score
    question_num = 0
    score = 0
    mode_frame.pack()
    main_frame.pack_forget()
    next_question()

# 生成問題
def generate_question():
    global current_question, current_choices, current_answer

    correct_index = random.randint(0, len(words) - 1)
    correct_word = words[correct_index]
    correct_translation = translations[correct_index]
    
    if selected_mode == 1:  # 中翻英
        current_question = f"翻譯成英文: {correct_translation}"
        current_choices = random.sample([word for word in words if word != correct_word], 2)
        current_choices.append(correct_word)
        random.shuffle(current_choices)
        current_answer = correct_word

    elif selected_mode == 2:  # 英翻中
        current_question = f"翻譯成中文: {correct_word}"
        current_choices = random.sample([trans for trans in translations if trans != correct_translation], 2)
        current_choices.append(correct_translation)
        random.shuffle(current_choices)
        current_answer = correct_translation

    elif selected_mode == 3:  # 克漏字
        sentence = sentences[correct_index].replace(correct_word, "____")
        current_question = f"請填入正確的單字: {sentence}"
        current_answer = correct_word

# 顯示下一題
def next_question():
    generate_question()
    show_question()

# 回到主選單
def main_menu():
    mode_frame.pack_forget()
    main_frame.pack()

# 介面元件
main_frame = tk.Frame(root)
main_frame.pack(pady=20)

welcome_label = tk.Label(main_frame, text="歡迎使用英文測試軟體！請選擇測試模式：", font=("Arial", 18))
welcome_label.pack(pady=10)

mode1_btn = tk.Button(main_frame, text="1. 中翻英", command=lambda: select_mode(1), width=30, height=2, font=("Arial", 14))
mode1_btn.pack(pady=5)

mode2_btn = tk.Button(main_frame, text="2. 英翻中", command=lambda: select_mode(2), width=30, height=2, font=("Arial", 14))
mode2_btn.pack(pady=5)

mode3_btn = tk.Button(main_frame, text="3. 克漏字", command=lambda: select_mode(3), width=30, height=2, font=("Arial", 14))
mode3_btn.pack(pady=5)

exit_btn = tk.Button(main_frame, text="退出程式", command=root.quit, width=30, height=2, font=("Arial", 14))
exit_btn.pack(pady=5)

# 顯示題目和選項
mode_frame = tk.Frame(root)

question_label = tk.Label(mode_frame, text="", font=("Arial", 16))
question_label.pack(pady=20)

options_frame = tk.Frame(mode_frame)
options_frame.pack(pady=10)

# 啟動介面
main_menu()
root.mainloop()
