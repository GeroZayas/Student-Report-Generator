# INSPIRED BY Sven from Coding Is Fun

# Report Generator [version 0.2] [Nov 6th, 2022]

import datetime
from pathlib import Path


import PySimpleGUI as sg
from docxtpl import DocxTemplate


# THIS CODE IS ADDED TO BE ABLE TO USE IMAGES WHEN MADE EXECUTABLE
import os
import sys


# To understand this solution applying this function of resource_path go to->
# https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
# specifically to "A clear and unambiguous guide"
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# ----------------------------------------------------------------


document_path = resource_path(Path(__file__).parent / "./model.docx")
doc = DocxTemplate(document_path)

today = datetime.datetime.today()

sg.theme("LightBrown3")

# Fonts
FONT = "Consolas 12"
TEXT_SIZE = (22, 1)
TOOLTIP_EXAM_MARKS = "'--': no se evaluó en la prueba | 'NA': no asistió en la prueba"


layout = [
    [
        sg.Text("Student:", TEXT_SIZE),
        sg.Input(key="STUDENT", do_not_clear=False, size=(30, 30), font=FONT),
    ],
    [
        sg.Text("Level:", TEXT_SIZE),
        sg.Input(key="LEVEL", do_not_clear=True, size=(30, 30), font=FONT),
    ],
    [
        sg.Text("Teacher:", TEXT_SIZE),
        sg.Input(key="TEACHER", do_not_clear=True, size=(30, 30), font=FONT),
    ],
    [
        sg.Text("Periodo:", TEXT_SIZE),
        sg.Input(
            key="PERIOD",
            do_not_clear=True,
            size=(30, 30),
            font=FONT,
            default_text="1er Trimestre",
        ),
    ],
    # --------------- GENERAL MARKS  ---------------
    [
        sg.Text("Asistencia:", TEXT_SIZE),
        sg.Input(key="ASISTENCIA", do_not_clear=False, size=(10, 10), font=FONT),
    ],
    [
        sg.Text("Asimilación de material nuevo:", TEXT_SIZE),
        sg.Input(key="ASIMILACION", do_not_clear=False, size=(10, 10), font=FONT),
    ],
    [
        sg.Text("Aprendizaje/Deberes:", TEXT_SIZE),
        sg.Input(key="APRENDIZAJE", do_not_clear=False, size=(10, 10), font=FONT),
    ],
    [
        sg.Text("Participación en clase/Interés:", TEXT_SIZE),
        sg.Input(key="PARTICIPACION", do_not_clear=False, size=(10, 10), font=FONT),
    ],
    [
        sg.Text("Comportamiento:", TEXT_SIZE),
        sg.Input(key="COMPORTAMIENTO", do_not_clear=False, size=(10, 10), font=FONT),
    ],
    [
        sg.Text("Progreso durante del trimestre:", TEXT_SIZE),
        sg.Input(key="PROGRESO", do_not_clear=False, size=(10, 10), font=FONT),
    ],
    [
        sg.Text("Prueba:", TEXT_SIZE),
        sg.Input(
            key="PRUEBA",
            do_not_clear=True,
            size=(30, 30),
            font=FONT,
            default_text="Trimestral",
        ),
    ],
    # --------------- EXAM MARKS ---------------
    [
        sg.Text("Listening:", TEXT_SIZE),
        sg.Input(
            key="LISTENING",
            do_not_clear=False,
            size=(10, 10),
            font=FONT,
            tooltip=TOOLTIP_EXAM_MARKS,
        ),
    ],
    [
        sg.Text("Reading and Use of Language:", TEXT_SIZE),
        sg.Input(
            key="READING_USE_LANGUAGE",
            do_not_clear=False,
            size=(10, 10),
            font=FONT,
            tooltip=TOOLTIP_EXAM_MARKS,
        ),
    ],
    [
        sg.Text("Writing:", TEXT_SIZE),
        sg.Input(
            key="WRITING",
            do_not_clear=False,
            size=(10, 10),
            font=FONT,
            tooltip=TOOLTIP_EXAM_MARKS,
        ),
    ],
    [
        sg.Text("Speaking:", TEXT_SIZE),
        sg.Input(
            key="SPEAKING",
            do_not_clear=False,
            size=(10, 10),
            font=FONT,
            tooltip=TOOLTIP_EXAM_MARKS,
        ),
    ],
    # COMENTARIO is multiline
    # The length of COMENTARIO shoul be more or less 460 characters, including spaces
    [
        sg.Text("Comentario:"),
        sg.Multiline(
            key="COMENTARIO",
            font=FONT,
            do_not_clear=False,
            size=(60, 7),
            autoscroll=False,
        ),
    ],
    # --------------- END OF NOTAS ---------------
    # CREATE BUTTON
    [
        sg.Text("Despedida:", TEXT_SIZE),
        sg.Input(
            key="DESPEDIDA",
            do_not_clear=True,
            size=(30, 30),
            font=FONT,
            default_text="¡Felices Vacaciones!",
        ),
    ],
    [
        sg.Button(
            key="GENERATE",
            image_filename=resource_path("button.png"),
            button_color="white",
            auto_size_button=True,
            border_width=0,
        ),
        sg.Exit(),
        sg.Text("Made by Gero Zayas", background_color="Black", text_color="Gold"),
    ],
]

window = sg.Window(
    "Futur Idiomes - Report Generator",
    layout,
    element_justification="left",
    icon=resource_path("futur_logo.ico"),
    resizable=False,
)


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "GENERATE":

        try:
            values["TOTAL"] = (
                float(values["LISTENING"])
                + float(values["READING_USE_LANGUAGE"])
                + float(values["WRITING"])
                + float(values["SPEAKING"])
            ) / 4

        except Exception:
            values["TOTAL"] = "..."

        # Render the template, save new word document & inform user
        doc.render(values)

        doc.save(f"./{values['STUDENT']}-{values['LEVEL']}.docx")
        sg.popup("File has been saved!")

window.close()
