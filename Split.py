import re

#my_string = "abc123"
#M6392ASCOT NERO-SENZA ETICHETTA
my_string = "M6392ASCOT NERO-SENZA ETICHETTA"
match = re.search(r"(\D+)(\d+)$", my_string)

if match:
    first_part = match.group(1)
    second_part = match.group(2)
    print(first_part, second_part)