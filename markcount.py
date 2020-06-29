import openpyxl

from answerKey import answerKey


def getData(sheet):
    """
    Extracts data from origin file, parsing students’ information and exam
    answers. Returns it as two tuples per student: one containing their
    information, and another with their answers.
    Inputs: sheet, an Excel sheet object.
    Returns: tuples of various data types.
    """
    info, answers = [], []
    for rowNum in range(2, sheet.max_row + 1):
        infoTemp, answersTemp = [], []
        for colNum in range(2, 6):
            infoTemp.append(sheet.cell(row=rowNum, column=colNum).value)
        for i in range(4, 6):
            infoTemp.insert(4, infoTemp[0])
            del infoTemp[0]
        info.append(tuple(infoTemp))
        for colNum in range(7, 78):
            content = sheet.cell(row=rowNum, column=colNum).value
            # Countermeasure for error in exam. Remove block if exam is fixed.
            try:
                answersTemp.append(content[0] if "," not in content else None)
            except TypeError:
                answersTemp.append(None)
        answers.append(tuple(answersTemp))
    return tuple(info), tuple(answers)


def checkAnswers(answers, answerKey):
    """
    Checks students’ answers against answer key and returns a tuple of scores
    per student.
    Inputs: answers, a tuple of tuples of strings.
            answerKey, a tuple of strings.
    Returns: a tuple of tuples of integers.
    """
    def checkSection(current, start, end):
        """
        Check answers in a section.
        Inputs: current, a list of strings.
                start, an integer.
                end, an integer.
        Returns: an integer.
        """
        counter = 0
        for i in range(start, end):
            if current[i] == answerKey[i]:
                counter += 1
        return counter

    final = []
    for elem in answers:
        results = [
            checkSection(elem, 0, 12),
            checkSection(elem, 12, 24),
            checkSection(elem, 24, 36),
            checkSection(elem, 36, 47),
            checkSection(elem, 47, 57),
            checkSection(elem, 57, 71)
        ]
        final.append(tuple(results))
    return tuple(final)


def writeResults():
    """
    Write results per section to a new Excel file.
    Returns: nothing.
    """
    wb = openpyxl.load_workbook("marks.xlsx")
    sheet = wb.active
    info, answers = getData(sheet)
    sections = checkAnswers(answers, answerKey)

    wb = openpyxl.Workbook()
    sheet = wb.active
    values = ("First Name", "Last Name", "Email", "Section 1", "Section 2",
              "Section 3", "Section 4", "Section 5", "Section 6", "Old score",
              "New score")
    for colNum in range(len(values)):
        sheet.cell(row=1, column=colNum + 1).value = values[colNum]
    for rowNum in range(2, len(info) + 2):
        for colNum in range(1, 4):
            sheet.cell(
                row=rowNum, column=colNum
            ).value = info[rowNum - 2][colNum - 1]
        for colNum in range(4, 10):
            sheet.cell(
                row=rowNum, column=colNum
            ).value = sections[rowNum - 2][colNum - 4]
        sheet.cell(row=rowNum, column=10).value = info[rowNum - 2][3]
        sheet.cell(row=rowNum, column=11).value = sum(sections[rowNum - 2])
    wb.save("results.xlsx")


writeResults()
