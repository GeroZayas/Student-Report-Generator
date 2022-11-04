import datetime
from pathlib import Path

import PySimpleGUI as sg
from docxtpl import DocxTemplate

document_path = Path(__file__).parent / "./model.docx"
doc = DocxTemplate(document_path)

today = datetime.datetime.today()

sg.theme("BrightColors")

# Fonts
FONT = "Consolas 12"
TEXT_SIZE = (22, 1)


layout = [
    # [sg.Image(source="./futur_logo.jpg")],
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
        # FIXME:
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
        ),
    ],
    [
        sg.Text("Reading and Use of Language:", TEXT_SIZE),
        sg.Input(
            key="READING_USE_LANGUAGE",
            do_not_clear=False,
            size=(10, 10),
            font=FONT,
        ),
    ],
    [
        sg.Text("Writing:", TEXT_SIZE),
        sg.Input(
            key="WRITING",
            do_not_clear=False,
            size=(10, 10),
            font=FONT,
        ),
    ],
    [
        sg.Text("Speaking:", TEXT_SIZE),
        sg.Input(
            key="SPEAKING",
            do_not_clear=False,
            size=(10, 10),
            font=FONT,
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
    # PUT the length of what is written in COMENTARIO
    # [sg.Text(f"Length: {length_of_comment}")],
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
            image_filename="./3 GERO - Create GUI - copia/Report_Generator/button.png",
            button_color="white",
            auto_size_button=True,
            border_width=0,
        ),
        sg.Exit(button_color="white"),
    ],
]

window = sg.Window(
    "Futur Idiomes - Report Generator",
    layout,
    element_justification="left",
    icon="./3 GERO - Create GUI - copia/Report_Generator/futur_logo.ico",
    resizable=False,
)


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "GENERATE":
        # print(event, values)

        # Calculate total

        # si una destreza se asigna con --, significa que no se evaluó en la prueba
        # si una destreza se asigna con NA, significa que el alumno no asistió en la prueba

        for value in values:
            if values[value] == "--" or values[value] == "NA":
                values[value] = 0

        # FIXME: print in final report -- or NA not 0

        values["TOTAL"] = (
            float(values["LISTENING"])
            + float(values["READING_USE_LANGUAGE"])
            + float(values["WRITING"])
            + float(values["SPEAKING"])
        ) / 4

        # Render the template, save new word document & inform user
        doc.render(values)
        output_path = (
            Path(__file__).parent / f"{values['STUDENT']}-{values['LEVEL']}.docx"
        )
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")

window.close()
