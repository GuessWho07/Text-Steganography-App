from docx import Document
from math import floor

def covert_string_to_bits(text):
    bits = ""
    for char in text:
        ascii_code = ord(char)
        binary_code = bin(ascii_code)[2:].zfill(8)  # Ensure it's 8 bits long
        bits += binary_code
    return bits

def convert_bits_to_string(bits):
    result = ""
    for i in range(0, len(bits), 8):
        byte = bits[i:i+8]
        decimal_value = int(byte, 2)
        result += chr(decimal_value)
    return result

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

    return potential

def get_modified_color(original_color, stegano_value, number_of_bytes):
    print(bin(original_color))
    print(bin(stegano_value))
    binary_oc = bin(original_color)[2:number_of_bytes+2]
    binary_sv = bin(stegano_value)[2:]
    binary_sv = bin(0)[2:]*(4-len(binary_sv))+binary_sv
    print("COLORS: ",binary_oc, binary_sv)
    print("NEW_COLOR: {0}{1}".format(binary_oc,binary_sv))

    return binary_oc+binary_sv

def hide_message(message,document,depth):
    # Zamień message na ciąg bitów
    message_bytes = covert_string_to_bits(message)
    print("Encoded message: ",message_bytes)
    
    # Zakoduj to na ostatnich 4 bitach koloru R pierwszej spacji
    binary_depth = bin(depth)[2:]
    print(binary_depth)
    print(get_modified_color(255,depth,4))
    #TODO: Faktycznie zmień kolor spacji, wgle trzeba jakoś trackować spacje które zmieniamy i ich wartości RGB

    # Podziel strumień bitów na supergrupy po 3 grupy po X bitów, ew. dodaj padding
    divided_message = []
    for i in range(0, len(message_bytes),3*depth):
        divided_message.append(message_bytes[i:i+(3*depth)])

    print(divided_message)
    print(convert_bits_to_string(''.join(divided_message)))

    for i in range(len(divided_message)):
        # Weź spację
        # Podziel supergrupę na 3
        # Zmień wartości RGB
        continue

    # Weź spację, weź jej kolor, dodaj do niego tajną wartość / Na razie zakładamy że jest czarna
    # ^^^ Powtórz ile razy trzeba
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
            full_path="C:\\"+full_path

            while True:
                while True:
                    depth = int(input("Please provide depth of the message (range from 1 to 8). Lower values will make the font color less distinghushible from color black, but also less characters will be coded within one space\nYour choice (1-8): "))
                    if depth in range (1,9):
                        break
                    else:
                        print("Invalid value, must be in range (1,8)")
                        continue
                
                # Calculate max length of message that can be hidden in this document
                document = Document(full_path)
                potential = calculate_doc_potential(document,depth)
                message = input("Please, provide message you want to hide in your document\nMessage: ")
                if len(message)>potential:
                    print("Your message is too long for your doc, please pick other parameters...")
                    continue
                else:
                    break
            
            hide_message(message,document,depth)

        elif action == 2:
            full_path = input("Please, provide full PATH to the Document you want to extract message from...\n")
        else:
            print("Wrong action has been chosen...")
            continue


        
if __name__ == "__main__":
    main()
