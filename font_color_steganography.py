from docx import Document
from docx.shared import RGBColor
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
    # print(bin(original_color))
    # print(bin(stegano_value))
    binary_oc = bin(original_color)[2:10-number_of_bytes]
    binary_sv = bin(stegano_value)[2:]
    binary_sv = bin(0)[2:]*(4-len(binary_sv))+binary_sv
    # print("COLORS: ",binary_oc, binary_sv)
    print("NEW_COLOR: {0}{1}".format(binary_oc,binary_sv))

    return binary_oc+binary_sv

def get_modified_color_from_bits(original_color, stegano_value_bits, number_of_bytes):
    # print(bin(original_color))
    # print(bin(stegano_value))
    binary_oc = bin(original_color)[2:10-number_of_bytes]
    binary_sv = stegano_value_bits
    binary_sv = bin(0)[2:]*(4-len(binary_sv))+binary_sv
    # print("COLORS: ",binary_oc, binary_sv)
    print("NEW_COLOR: {0}{1}".format(binary_oc,binary_sv))

    return binary_oc+binary_sv

def change_space_color(doc_path, new_doc_path):
    # Load the document
    doc = Document(doc_path)

    # Define the RGB color for the font
    new_color = RGBColor(255, 255, 255)  # This is red, you can change it to any RGB value you want

    # Iterate through paragraphs
    for para in doc.paragraphs:
        for run in para.runs:
            # Check if the run contains a space
            if ' ' in run.text:
                # Split the run by spaces
                parts = run.text.split(' ')
                new_run = para.add_run()  # Create a new run
                new_run.style = run.style  # Copy the style from the original run

                # Iterate through parts and add them with the new color
                for part in parts:
                    if part == '':
                        new_run.add_text(' ')
                    else:
                        new_run.add_text(part)
                        new_run.font.color.rgb = new_color

                # Remove the original run (optional)
                para.runs.remove(run)

    # Save the modified document
    doc.save(new_doc_path)


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

    spaces_rgbs = []
    spaces_rgbs.append(get_modified_color(255,depth,4))

    for supergroup in divided_message:
        R = supergroup[:depth]
        G = supergroup[depth:2*depth]
        B = supergroup[2*depth:]
        print(supergroup,"---->",R,G,B)
        print('Stegano: ',get_modified_color_from_bits(255,R,depth),get_modified_color_from_bits(255,G,depth),get_modified_color_from_bits(255,B,depth))
        spaces_rgbs.append(get_modified_color_from_bits(255,R,depth))
        spaces_rgbs.append(get_modified_color_from_bits(255,G,depth))
        spaces_rgbs.append(get_modified_color_from_bits(255,B,depth))
        # Weź spację
        # Podziel supergrupę na 3
        # Zmień wartości RGB

    print(spaces_rgbs)
    return spaces_rgbs
    # Weź spację, weź jej kolor, dodaj do niego tajną wartość / Na razie zakładamy że jest czarna
    # ^^^ Powtórz ile razy trzeba

def show_message(hidden_message_list):
    ## ODWRÓC PROCES - DLA TESTÓW
    print(hidden_message_list[0][4:])
    print(int(hidden_message_list[0][4:],2))
    dec_depth = int(hidden_message_list[0][4:],2)
    hidden_message_list.pop(0)
    decoding_string = ""
    for element in hidden_message_list:
        print("ELEMENT:",element)
        decoding = element[8-dec_depth:]
        print("DECODING:",decoding)
        decoding_string += decoding

    print("STREAM:",decoding_string)
    decoding_list = [decoding_string[i:i+8] for i in range(0, len(decoding_string), 8)]
    decrypted_message = ""
    for letter in decoding_list:
        decrypted_message+=chr(int(letter,2))

    print(decrypted_message)
    return decrypted_message

def main():
    while True:
        print("Please, choose action")
        action = int(input("1 - Hide message in document\n2 - Extract message from document\n Your choice...: "))
        if action == 1:
            full_path = input("Please, provide full PATH to the Document you want to hide message in...\n")
            # full_path="C:\\"+full_path

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
            

            print("For presentation purposes, now the message will be unveiled...")
            hidden_message = hide_message(message,document,depth)
            encrypted_message = show_message(hidden_message)

            print("Original message:",message)
            print("Hidden message (Each element of the list defines one of RBG components of Space color):",hidden_message)
            print("Unveiled message:",encrypted_message)

        elif action == 2:
            full_path = input("Please, provide full PATH to the Document you want to extract message from...\n")
        else:
            print("Wrong action has been chosen...")
            continue


        
if __name__ == "__main__":
    main()
