def covert_string_to_bytes(text):
    print("TODO: Write function")

def convert_bytes_to_string(bytes):
    print("TODO: Write function")

def get_modified_color(original_color, stegano_value, number_of_bytes):
    print("TODO: Add stegano_value to the original color on the defined number of bytes")

def hide_message(message):
    # Zamień message na ciąg bitów
    # Zdefiniuj na ilu bitach przechowujemy wiadomość
    # Zakoduj to na ostatnich 3 bitach pierwszej spacji
    # Podziel strumień bitów na supergrupy po 3 grupy po X bitów, ew. dodaj padding
    # Weź spację, weź jej kolor, dodaj do niego tajną wartość
    # ^^^ Powtórz ile razy trzeba
    # Dodaj na początku 3 bitową wartość, definiującą na ilu bitach chowamy wiadomość
    pass

def show_message(message):
    # Weź pierwszą spację, przeczytaj na ilu bitach chowamy wiadomość
    pass
