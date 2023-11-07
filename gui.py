import tkinter as tk
from tkinter import filedialog
import space

def encrypt_button_click():
    text_file = filedialog.askopenfilename(title="Select Text File")
    word_file = filedialog.askopenfilename(title="Select Word File", defaultextension=".docx")
    if text_file and word_file:
        space.encrypt(text_file, word_file)
        status_label.config(text="Encryption Complete")

def decrypt_button_click():
    word_file = filedialog.askopenfilename(title="Select Word File")
    if word_file:
        space.decrypt(word_file)
        status_label.config(text="Decryption Complete")

root = tk.Tk()
root.title("Text Encryption and Decryption")

encrypt_button = tk.Button(root, text="Encrypt Text", command=encrypt_button_click)
decrypt_button = tk.Button(root, text="Decrypt Text", command=decrypt_button_click)
status_label = tk.Label(root, text="", padx=10)

encrypt_button.pack(pady=10)
decrypt_button.pack(pady=10)
status_label.pack()

root.mainloop()