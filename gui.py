import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog, messagebox
import space
import font_color_steganography as fc
from docx import Document

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

def encrypt_button_click_fontcolor():
    # Wybierz plik
    word_file = filedialog.askopenfilename(title="Select Word File", defaultextension=".docx")
    doc = Document(word_file)
    # Wybierz głębokość
    depth = simpledialog.askinteger("Input","Enter depth", initialvalue=4, minvalue=1, maxvalue=8)
    # Pokaż ile znaków można zakodować
    max_chars = fc.calculate_doc_potential(doc,depth)
    message_to_messagebox = "Maximum number of characters that can be hidden is {}".format(max_chars)
    messagebox.showinfo("Maximum number of characters",message_to_messagebox)
    # Wpisz tajną wiadomość
    message = simpledialog.askstring("Input","Enter Message")
    # Zakoduj
    fc.hide_message(message,word_file,depth)
    

def decrypt_button_click_fontcolor():
    # Wybierz plik
    word_file = filedialog.askopenfilename(title="Select Word File", defaultextension=".docx")
    # Pokaż ukrytą wiadomość
    decrypted_message = fc.show_message_from_file(word_file)
    message = "Decrypted message: {}".format(decrypted_message)
    messagebox.showinfo("Maximum number of characters",message)

root = tk.Tk()
root.title("Text Encryption and Decryption")

encrypt_button = tk.Button(root, text="Encrypt Text", command=encrypt_button_click)
decrypt_button = tk.Button(root, text="Decrypt Text", command=decrypt_button_click)
status_label = tk.Label(root, text="", padx=10)
encrypt_button_fc = tk.Button(root, text="Encrypt Text - Font Color", command=encrypt_button_click_fontcolor)
decrypt_button_fc = tk.Button(root, text="Decrypt Text - Font Color", command=decrypt_button_click_fontcolor)

encrypt_button.pack(pady=10)
decrypt_button.pack(pady=10)
encrypt_button_fc.pack(pady=10)
decrypt_button_fc.pack(pady=10)
status_label.pack()

root.mainloop()