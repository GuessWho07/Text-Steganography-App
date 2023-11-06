from docx import Document
from docx.shared import RGBColor
from math import floor
from lxml import etree

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
    binary_sv = bin(0)[2:]*(number_of_bytes-len(binary_sv))+binary_sv
    # print("COLORS: ",binary_oc, binary_sv)
    print("NEW_COLOR: {0}{1}".format(binary_oc,binary_sv))

    return binary_oc+binary_sv

def get_modified_color_from_bits(original_color, stegano_value_bits, number_of_bytes):
    # print(bin(original_color))
    # print(bin(stegano_value))
    binary_oc = bin(original_color)[2:10-number_of_bytes]
    binary_sv = stegano_value_bits
    binary_sv = bin(0)[2:]*(number_of_bytes-len(binary_sv))+binary_sv
    # print("COLORS: ",binary_oc, binary_sv)
    print("NEW_COLOR: {0}{1}".format(binary_oc,binary_sv))

    return binary_oc+binary_sv

# def split_run_by_char(run):
#     new_runs = []
#     for char in run.text:
#         new_run = run._element.makeelement(run._r.tag)
#         new_run.text = char
#         new_run._r.clear_content()
#         new_run._r.append(new_run._element)
#         new_runs.append(new_run)
#     return new_runs

def print_run_style(run):
    print("Run boldness:",run.bold)
    print("Run italic:",run.italic)
    print("Run underline:",run.underline)
    print("Run font:",run.font.name)
    print("Run fontsize:",run.font.size)
    print("Run color:",run.font.color.rgb)

def copy_run_style(source_run, target_run):
    # print_run_style(source_run)
    # print_run_style(target_run)
    target_run.bold = source_run.bold
    target_run.italic = source_run.italic
    target_run.underline = source_run.underline
    target_run.font.name = source_run.font.name
    target_run.font.size = source_run.font.size
    target_run.font.color.rgb = source_run.font.color.rgb

def change_space_color(doc_path, new_doc_path,colors):
    doc = Document(doc_path)
    doc_new = Document()

    for triplet in colors:
        while True:
            if len(triplet)<3:
                triplet.append(0)
            elif len(triplet) == 3:
                break

    print("Final colors list:",colors)
    color_index = 0  # Initialize color index

    for para_index, paragraph in enumerate(doc.paragraphs):
        doc_new.add_paragraph()
        for run_index,run in enumerate(paragraph.runs):
            print("Run text:",run.text)

            space_count = 0
            space_pos = []
            runs_list = []

            for i in range(len(run.text)):
                if run.text[i] == ' ':
                    space_count += 1
                    space_pos.append(i)

            run_text = run.text
            print("Para runs",paragraph.runs,run_index)
            paragraph.runs.pop(run_index)
            print("Para runs after del",paragraph.runs)
            pointer = 0
            for position in space_pos:
                runs_list.append(paragraph.add_run(run_text[pointer:position]))
                runs_list.append(paragraph.add_run(run_text[position]))
                pointer = position + 1

            runs_list.append(paragraph.add_run(run_text[pointer:]))

            for i, run in enumerate(runs_list):
                print("ID, run",i,run.text)

            for run_part_index, run_part in enumerate(runs_list):
                # paragraph.runs.insert(run_index+run_part_index,run_part)
                doc_new.paragraphs[para_index].add_run(run_part.text)
                print_run_style(run_part)
                print_run_style(doc_new.paragraphs[para_index].runs[run_part_index])
                copy_run_style(run_part,doc_new.paragraphs[para_index].runs[run_part_index])

            for index,run in enumerate(doc_new.paragraphs[para_index].runs):
                if run.text == ' ' and color_index < len(colors):
                    print("Colors, color_index:",colors,color_index)
                    r, g, b = colors[color_index]
                    run.font.color.rgb = RGBColor(r,g,b)
                    color_index += 1


            # for i in range(len(run.text)):
            #     if run.text[i] == ' ' and color_index < len(colors):
            #         # Get the RGB values for the current space
            #         r, g, b = colors[color_index]
            #         print("Current colors:",colors[color_index])
            #         print("Extracted Values:",r,g,b)
            #         print("Current char:",run.text[i],"Previous:",run.text[i-1],"Next:",run.text[i+1],"Broader:",run.text[i-2:i+2])

            #         text = run.text
            #         paragraph.runs.pop(run_index)

            #         run_b = paragraph.add_run(text[:i])
            #         run_d = paragraph.add_run(text[i])
            #         run_a = paragraph.add_run(text[i+1:])

            #         print("Run B text:",run_b.text,"Run D text:",run_d.text,"Run A text:",run_a.text)
                    
            #         run_d.font.color.rgb = RGBColor(r,g,b)
            #         paragraph.runs.insert(run_index-1,run_b)
            #         paragraph.runs.insert(run_index,run_d)
            #         paragraph.runs.insert(run_index+1,run_a)

            #         color_index += 1  # Move to the next color

    

    doc_new.save(new_doc_path)


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

    print("Spaces RGBs BIN",spaces_rgbs)
    spaces_rgbs_dec = [int(binary, 2) for binary in spaces_rgbs]
    print("Spaces RGBs DEC",spaces_rgbs_dec)
    spaces_rgbs_dec_triplet = [spaces_rgbs_dec[i:i+3] for i in range(0, len(spaces_rgbs_dec), 3)]
    print("Spaces RGBs DEC TRIPLET",spaces_rgbs_dec_triplet)
    
    change_space_color(document,"C:\\Users\paluc\OneDrive\Dokumenty\out.docx",spaces_rgbs_dec_triplet)

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
            # full_path = input("Please, provide full PATH to the Document you want to hide message in...\n")
            full_path = "C:\\Users\paluc\OneDrive\Dokumenty\\a.docx"
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
            hidden_message = hide_message(message,full_path,depth)
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
