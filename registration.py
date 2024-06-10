import cv2
import os
import tkinter as tk
from tkinter import simpledialog, messagebox, ttk
from PIL import Image, ImageTk

# Функция для сохранения фото
def take_photo(username):
    # cap = cv2.VideoCapture(0)
    
    if not cap.isOpened():
        messagebox.showerror("Error", "Could not open video device")
        return
    
    ret, frame = cap.read()
    if not ret:
        messagebox.showerror("Error", "Failed to grab frame")
        return
    
    filename = f"dataset/{username.replace(' ', '_')}.jpg"
    cv2.imwrite(filename, frame)
    cap.release()
    messagebox.showinfo("Success", f"Photo saved as {filename}")

# Функция для обновления видеопотока
def update_frame():
    ret, frame = cap.read()
    if not ret:
        return
    
    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    img = Image.fromarray(frame)
    imgtk = ImageTk.PhotoImage(image=img)
    video_label.imgtk = imgtk
    video_label.configure(image=imgtk)
    video_label.after(10, update_frame)

# Функция для обработки кнопки "Сделать снимок"
def capture_photo():
    username = username_entry.get()
    if not username:
        messagebox.showerror("Error", "Please enter your full name")
        return
    take_photo(username)

def retake_photo():
    global cap
    cap = cv2.VideoCapture(0)
    update_frame()
    return


root = tk.Tk()
root.title("User Registration")
root.configure(background="#B2DFDB")

ttk.Style().configure(".",  font="helvetica 13", foreground="#004D40", padding=8, background="#B2DFDB")

# Поле для ввода ФИ
ttk.Label(root, text="Enter your full name:").pack(pady=10)
username_entry = ttk.Entry(root, width=50)  
username_entry.pack(pady=10)

button_frame = ttk.Frame(root)

# Кнопка для снимка
capture_button = ttk.Button(button_frame, text="Capture Photo", command=capture_photo)
capture_button.pack(side="left", pady=10, padx=10)

retake_button = ttk.Button(button_frame, text="Retake Photo", command=retake_photo)
retake_button.pack(side="right", pady=10, padx=10)

button_frame.pack()

# Виджет для видео
video_label = ttk.Label(root)
video_label.pack()

# Запуск видеокамеры
cap = cv2.VideoCapture(0)
update_frame()

# Запуск интерфейса
root.mainloop()

# Освобождение ресурсов
# cap.release()
cv2.destroyAllWindows()
