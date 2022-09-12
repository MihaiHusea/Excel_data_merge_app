def range_char(start, stop):
    return (chr(n) for n in range(ord(start), ord(stop) + 1))

for character in range_char("A", "D"):
    print(character)

lista=[str(character)+'2'for character in range_char("A", "D") ]
print(lista)