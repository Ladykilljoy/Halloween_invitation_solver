from base64 import *
from oletools.olevba import *
import re


# Imports the VBA source code and extracts the obfuscated strings that includes the flag, and arranges them in pairs
def get_obfuscated_vba_strings(path):
    obfuscated_string_list = []
    vbaparser = VBA_Parser(path)
    results = vbaparser.analyze_macros()
    if vbaparser.detect_vba_macros():
        print("VBA Macro found")
        print("__________________________")
        for filename, stream_path, vba_filename, vba_code in vbaparser.extract_macros():
            for line in vba_code.splitlines():
                obfuscated_string = re.search(r'\"(\d{2,})\".*\"(\d{2,})\"', line)
                if obfuscated_string:
                    obfuscated_string_list.extend([obfuscated_string.group(1), obfuscated_string.group(2)])
    else:
        print("No VBA Macro found")
    obfuscated_string_pairs = [obfuscated_string_list[x:x + 2] for x in range(0, len(obfuscated_string_list), 2)]
    return obfuscated_string_pairs

# decodes the string pairs we got from
def solve_ctf(path_of_htb_file):
    string_pairs = get_obfuscated_vba_strings(path_of_htb_file)
    solve = ""

    for x in string_pairs:
        solve = solve + int_to_str(hex_to_int(x[0]) + hex_to_int(x[1]))

    print("The full powershell payload is: " + b64decode(solve).decode('utf-16'))
    flag = re.search(r"(HTB.+)", b64decode(solve).decode('utf-16'))
    print("The flag is: " + flag.group(1))


def hex_to_int(text):
    rtn = ""
    for i in range(0, len(text), 2):
        rtn = rtn + chr(int(text[i:i + 2], 16))
    else:
        return rtn


def int_to_str(text):
    rtn = ""
    for x in text.split(" "):
        rtn = rtn + chr(int(x))
    else:
        return rtn


if __name__ == '__main__':
    input_path = input("Enter full path of \"invitation.docm\": ")
    solve_ctf(input_path)
