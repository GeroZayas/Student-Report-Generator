import datetime
from pathlib import Path

import PySimpleGUI as sg
from docxtpl import DocxTemplate

document_path = Path(__file__).parent / "./model.docx"
doc = DocxTemplate(document_path)

today = datetime.datetime.today()

sg.theme("BrightColors")


layout = [
    # [sg.Image(source="./futur_logo.jpg")],
    [
        sg.Text("Student:"),
        sg.Input(key="STUDENT", do_not_clear=False, size=(30, 30), font="Consolas 14"),
    ],
    [
        sg.Text("Level:"),
        sg.Input(key="LEVEL", do_not_clear=True, size=(30, 30), font="Consolas 14"),
    ],
    [
        sg.Text("Teacher:"),
        sg.Input(key="TEACHER", do_not_clear=True, size=(30, 30), font="Consolas 14"),
    ],
    [
        sg.Text("Periodo:"),
        sg.Input(key="PERIOD", do_not_clear=True, size=(30, 30), font="Consolas 14"),
    ],
    [
        sg.Text("Asistencia:"),
        sg.Input(
            key="ASISTENCIA", do_not_clear=False, size=(10, 10), font="Consolas 14"
        ),
    ],
    [
        sg.Text("Asimilación de material nuevo:"),
        sg.Input(
            key="ASIMILACION", do_not_clear=False, size=(10, 10), font="Consolas 14"
        ),
    ],
    [
        sg.Text("Aprendizaje/Deberes:"),
        sg.Input(
            key="APRENDIZAJE", do_not_clear=False, size=(10, 10), font="Consolas 14"
        ),
    ],
    [
        sg.Text("Participación en clase/Interés:"),
        sg.Input(
            key="PARTICIPACION", do_not_clear=False, size=(10, 10), font="Consolas 14"
        ),
    ],
    [
        sg.Text("Comportamiento:"),
        sg.Input(
            key="COMPORTAMIENTO", do_not_clear=False, size=(10, 10), font="Consolas 14"
        ),
    ],
    [
        sg.Text("Progreso durante del trimestre:"),
        sg.Input(key="PROGRESO", do_not_clear=False, size=(10, 10), font="Consolas 14"),
    ],
    [
        sg.Text("Prueba:"),
        sg.Input(key="PRUEBA", do_not_clear=True, size=(30, 30), font="Consolas 14"),
    ],
    # CREATE BUTTON
    [sg.Button("Create Report"), sg.Exit()],
]

window = sg.Window("Report Generator", layout, element_justification="right")

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Create Report":
        print(event, values)

        # Render the template, save new word document & inform user
        doc.render(values)
        output_path = Path(__file__).parent / f"{values['STUDENT']}-report.docx"
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")

window.close()
