from docx import Document
import os

def points_to_you(expected, result):
    ex = expected.split("-")
    if ex[0] == result[0] and ex[1] == result[1]:
        return 10
    elif result[0] > result[1] and ex[0] > ex[1] or\
        result[0] < result[1] and ex[0] < ex[1] or\
        result[0] == result[1] and ex[0] == ex[1]:
        return 6
    elif ex[0] == result[0] and ex[1] != result[1] or\
        ex[0] != result[0] and ex[1] == result[1]:
        return 3
    return 0

def printexceltemplate(filename):
    wordDoc = Document('1eKOOK/' + filename)
    iterator = 0
    j = 0
    points = 0
    for table in wordDoc.tables:
        for row in table.rows:
            # print(row.cells[0].text)
            if iterator == 0:
                hole_block = row.cells[0].text
                naam = hole_block.replace("Naam en mailadres: ", '')
                naam = naam.rstrip()
                print(naam, end=",")
            if iterator > 1 and iterator < 9:
                print(row.cells[2].text.strip(), end=",")
            if iterator == 9:
                print(row.cells[2].text.strip())
            iterator += 1
    return naam, points


lijst = os.listdir("1eKOOK")
ranking = []
for file in lijst: 
    printexceltemplate(file)
