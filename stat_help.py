
import pyperclip
import json

def convert_and_multiply(input_string, multiplier):
    elements = input_string.split("\t")
    
    processed_elements = []
    for element in elements:
        try:
            number = float(element.replace(" ", ""))
            multiplied_number = number * multiplier
            if multiplied_number.is_integer():
                multiplied_number = int(multiplied_number)
            processed_elements.append(multiplied_number)
        except ValueError:
            processed_elements.append(element)

    # Convert the list to JSON, without indenting
    json_string = json.dumps(processed_elements, ensure_ascii=False)
    # Instead of adding indents, manually format the string
    # to place each item on a new line and brackets on separate lines
    formatted_json_string = "[\n" + ",\n".join(json.dumps(item) for item in processed_elements) + "\n]"

    return formatted_json_string

# Main script
if __name__ == "__main__":
    input("Copy the cells from Excel to the clipboard and press Enter...")

    clipboard_text = pyperclip.paste()

    while True:
        try:
            multiplier = float(input("Please enter a multiplier: "))
            break
        except ValueError:
            print("Invalid input. Please enter a numeric value.")

    result = convert_and_multiply(clipboard_text, multiplier)

    print("\nResulting JSON Array:\n")
    print(result)

    pyperclip.copy(result)
    print("\nThe resulting JSON array has been copied to the clipboard.")
