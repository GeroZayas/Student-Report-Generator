# INSPIRED BY Sven from Coding Is Fun

# Report Generator [version 0.4] [Nov 7th, 2022]

import datetime
from pathlib import Path


import PySimpleGUI as sg
from docxtpl import DocxTemplate


# THIS CODE IS ADDED TO BE ABLE TO USE IMAGES WHEN MADE EXECUTABLE
import os
import sys


def main():

    # To understand this solution applying this function of resource_path go to->
    # https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
    # specifically to "A clear and unambiguous guide"
    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    # --------------------------------------------------------------------------

    document_path = resource_path(Path(__file__).parent / "./model.docx")
    doc = DocxTemplate(document_path)

    today = datetime.datetime.today()

    # -------------------------CHANGE THEMES ON THE FLY!------------------------
    selected_theme = "DefaultNoMoreNagging"
    sg.theme(selected_theme)

    # --------------------------------------------------------------------------

    new_folder = resource_path(Path(__file__).parent / "./00_GENERATED_REPORTS")

    if not os.path.exists(new_folder):
        os.makedirs(new_folder)

    # --------------------------------------------------------------------------

    # --------------------------------------------------------------------------

    # CONSTANTS
    TYPES_OF_PRUEBAS = ["Trimestral", "Final", "Final (Simulación de examen FCE)"]
    LEVELS = [
        "Young Learners",
        "Starters",
        "Movers",
        "Flyers",
        "KET",
        "PET",
        "FCE",
    ]
    PERIODOS = ["1er Trimestre", "2ndo Trimestre", "3er Trimestre"]

    # Fonts
    FONT = "Consolas 12"
    TEXT_SIZE = (22, 1)
    TOOLTIP_EXAM_MARKS = (
        "'--': no se evaluó en la prueba | 'NA': no asistió a la prueba"
    )

    # ----------------------------LAYOUT----------------------------

    layout = [
        [
            sg.Text("Student:", TEXT_SIZE),
            sg.Input(
                key="STUDENT",
                do_not_clear=False,
                size=(30, 30),
                font=FONT,
            ),
        ],
        [
            sg.Text("Level:", TEXT_SIZE),
            sg.Combo(key="LEVEL", values=LEVELS, size=(30, 30), font=FONT),
        ],
        [
            sg.Text("Teacher:", TEXT_SIZE),
            sg.Input(key="TEACHER", do_not_clear=True, size=(30, 30), font=FONT),
        ],
        [
            sg.Text("Periodo:", TEXT_SIZE),
            sg.Combo(
                values=PERIODOS,
                key="PERIOD",
                size=(30, 30),
                font=FONT,
                auto_size_text=True,
            ),
        ],
        [sg.HorizontalSeparator(color="white", pad=15)],
        # --------------- GENERAL MARKS  ---------------
        [
            sg.Text("Asistencia:", TEXT_SIZE),
            sg.Input(key="ASISTENCIA", do_not_clear=False, size=(10, 10), font=FONT),
            sg.Text("Asimilación de material nuevo:", TEXT_SIZE),
            sg.Input(key="ASIMILACION", do_not_clear=False, size=(10, 10), font=FONT),
        ],
        [
            sg.Text("Aprendizaje/Deberes:", TEXT_SIZE),
            sg.Input(key="APRENDIZAJE", do_not_clear=False, size=(10, 10), font=FONT),
            sg.Text("Participación en clase/Interés:", TEXT_SIZE),
            sg.Input(key="PARTICIPACION", do_not_clear=False, size=(10, 10), font=FONT),
        ],
        [
            sg.Text("Comportamiento:", TEXT_SIZE),
            sg.Input(
                key="COMPORTAMIENTO", do_not_clear=False, size=(10, 10), font=FONT
            ),
            sg.Text("Progreso durante del trimestre:", TEXT_SIZE),
            sg.Input(key="PROGRESO", do_not_clear=False, size=(10, 10), font=FONT),
        ],
        [sg.HorizontalSeparator(color="white", pad=15)],
        # --------------- EXAM MARKS ---------------
        [
            sg.Text("Prueba:", TEXT_SIZE),
            sg.Combo(
                values=TYPES_OF_PRUEBAS,
                key="PRUEBA",
                size=(30, 30),
                font=FONT,
                auto_size_text=True,
            ),
        ],
        [
            sg.Push(),
            sg.Text(
                "'--' -> no se evaluó en la prueba    |    'NA' -> no asistió a la prueba",
                text_color="grey",
            ),
            sg.Push(),
        ],
        [
            sg.Text("Listening:", TEXT_SIZE),
            sg.Input(
                key="LISTENING",
                do_not_clear=False,
                size=(10, 10),
                font=FONT,
                tooltip=TOOLTIP_EXAM_MARKS,
            ),
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
        [sg.HorizontalSeparator(color="white", pad=15)],
        [
            sg.Text("Comentario:"),
            sg.Multiline(
                key="COMENTARIO",
                font=FONT,
                do_not_clear=True,
                size=(60, 7),
                autoscroll=False,
            ),
        ],
        # --------------- END OF NOTAS ---------------
        # CREATE BUTTON
        [
            sg.Push(),
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
            sg.Push(),
            sg.Button(
                key="GENERATE",
                image_filename=resource_path("button.png"),
                button_color="white",
                auto_size_button=True,
                border_width=0,
            ),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Text("Made by Gero Zayas", background_color="Black", text_color="Gold"),
            sg.Push(),
        ],
    ]

    window = sg.Window(
        "Futur Idiomes - Report Generator (v0.4)",
        layout,
        element_justification="left",
        icon=resource_path("futur_logo.ico"),
        resizable=False,
    )

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "Exit":
            break

        if event == "select_theme":
            selected_theme = values["select_theme"]
            print(selected_theme)
            sg.theme(selected_theme)

        if event == "GENERATE":

            try:
                result_total = (
                    float(values["LISTENING"])
                    + float(values["READING_USE_LANGUAGE"])
                    + float(values["WRITING"])
                    + float(values["SPEAKING"])
                ) / 4
                values["TOTAL"] = round(result_total, 2)
            except Exception:
                values["TOTAL"] = "..."

            # Render the template, save new word document & inform user
            doc.render(values)

            doc.save(
                f"./00_GENERATED_REPORTS/{values['STUDENT']}-{values['LEVEL']}.docx"
            )
            sg.popup("File has been saved!")

    window.close()

    return None


if __name__ == "__main__":
    # This code won't run if this file is imported.
    main()
