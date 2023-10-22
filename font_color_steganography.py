from docx import Document
from math import floor

def covert_string_to_bytes(text):
    print("TODO: Write function")

def convert_bytes_to_string(bytes):
    print("TODO: Write function")

def calculate_doc_potential(document,depth):
    # Calculate how long message can be hidden in doc
    plaintext=""
    for para in document.paragraphs:
        plaintext += para.text + "\n"

    print("PLAINTEXT: ",plaintext)
    space_count=0

    for para in document.paragraphs:
        space_count += para.text.count(' ')

    print("SPACES: ",space_count)

    potential = floor((space_count*depth*3)/8)
    print("In this document, you can hide {0} characters".format(potential))

def get_modified_color(original_color, stegano_value, number_of_bytes):
    print("TODO: Add stegano_value to the original color on the defined number of bytes")

def hide_message(message,document,depth):
    # Zamień message na ciąg bitów
    # Zdefiniuj na ilu bitach przechowujemy wiadomość
    # Zakoduj to na ostatnich 3 bitach pierwszej spacji
    # Podziel strumień bitów na supergrupy po 3 grupy po X bitów, ew. dodaj padding
    # Weź spację, weź jej kolor, dodaj do niego tajną wartość
    # ^^^ Powtórz ile razy trzeba
    # Dodaj na początku 3 bitową wartość, definiującą na ilu bitach chowamy wiadomość
    pass

def show_message(document):
    # Weź pierwszą spację, przeczytaj na ilu bitach chowamy wiadomość
    pass

def main():
    while True:
        print("Please, choose action")
        action = int(input("1 - Hide message in document\n2 - Extract message from document\n Your choice...: "))
        if action == 1:
            full_path = input("Please, provide full PATH to the Document you want to hide message in...\n")

            while True:
                depth = int(input("Please provide depth of the message (range from 1 to 8). Lower values will make the font color less distinghushible from color black, but also less characters will be coded within one space\nYour choice (1-8): "))
                if depth in range (1,9):
                    break
                else:
                    print("Invalid value, must be in range (1,8)")
                    continue
            
            # Calculate max length of message that can be hidden in this document
            document = Document(full_path)
            calculate_doc_potential(document,depth)
            message = input("Please, provide message you want to hide in your document\nMessage: ")


        elif action == 2:
            full_path = input("Please, provide full PATH to the Document you want to extract message from...\n")
        else:
            print("Wrong action has been chosen...")
            continue


        
if __name__ == "__main__":
    main()
