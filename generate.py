import os
import openpyxl
import pptx

def generate(path_template, path_data, path_output, sheetname: str | None = None) -> None:
    wb = openpyxl.load_workbook(path_data)
    ws = wb.active
    if sheetname is not None:
        ws = wb[sheetname]
    
    keys = []
    for cell in ws[1]:
        if cell.value is None:
            break
        keys.append(cell.value)

    dir_output = os.path.splitext(path_output)[0]
    os.makedirs(dir_output, exist_ok=True)

    counter = 0
    while ws[counter + 2][0].value is not None:
        template = pptx.Presentation(path_template)
        for shape in template.slides[0].shapes:
            for i in range(len(keys)):
                if shape.text == keys[i]:
                    shape.text = ws[counter + 2][i]

        template.save(os.path.join(path_output, counter + '.pptx'))
        counter += 1