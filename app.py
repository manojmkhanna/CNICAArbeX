import logging

import gradio as gr
import pandas as pd
from google import genai
from pydantic import BaseModel

DEBUG = True
MAX_RESPONDENT_COUNT = 20
MAX_ADDRESS_HEADER_COUNT = 10
GEMINI_MODEL_NAME = "gemini-3-flash-preview"
GEMINI_PROMPT_PREFIX = "Split the following rows of names and addresses into columns such as Recipient Name/Entity Name, Address Line 1/Care of Name, Address Line 2, Address Line 3, District, State and PIN Code. " \
    "Add salutations like Mr., Mrs., Ms. or M/s. to Name and Care of Name if missing. " \
    "Add prefixes like s/o, d/o, f/o, m/o, h/o, w/o or c/o to Care of Name if missing. " \
    "Add punctuations to initials in Name and Care of Name if missing. " \
    "Correct spelling mistakes and punctuations in an Address if necessary. " \
    "Correct an incomplete Address if necessary. " \
    "Remove redundancy in an Address if necessary. " \
    "Convert everything to Proper case. " \
    "Do not ignore duplicate rows of names and addresses."
GEMINI_MAX_RESPONDENT_COUNT = 50


logging.basicConfig(
    format='%(asctime)s [%(levelname)s] %(message)s',
    level=logging.INFO
)


class Respondent(BaseModel):
    name: str
    address_line_1: str
    address_line_2: str
    address_line_3: str
    district: str
    state: str
    pin_code: str


class RespondentList(BaseModel):
    respondents: list[Respondent]


RespondentList.model_rebuild()


def original_excel_file_uploaded(original_excel_file_path):
    if original_excel_file_path is None:
        return [None] * (1 + MAX_RESPONDENT_COUNT + (MAX_RESPONDENT_COUNT * MAX_ADDRESS_HEADER_COUNT))

    original_excel_data_frame = pd.read_excel(original_excel_file_path)

    original_excel_column_headers = original_excel_data_frame.columns.tolist()

    name_header_dropdowns = []

    for i in range(MAX_RESPONDENT_COUNT):
        name_header_dropdown = gr.Dropdown(
            choices=original_excel_column_headers,
            value=original_excel_column_headers[0]
        )

        name_header_dropdowns.append(name_header_dropdown)

    address_header_dropdowns = []

    for i in range(MAX_RESPONDENT_COUNT):
        for j in range(MAX_ADDRESS_HEADER_COUNT):
            address_header_dropdown = gr.Dropdown(
                choices=original_excel_column_headers,
                value=original_excel_column_headers[0]
            )

            address_header_dropdowns.append(address_header_dropdown)

    return [original_excel_data_frame] + name_header_dropdowns + address_header_dropdowns


def respondent_slider_changed(respondent_count):
    respondent_tabs = []

    for i in range(MAX_RESPONDENT_COUNT):
        respondent_tab = gr.Tab(
            visible=i < respondent_count
        )

        respondent_tabs.append(respondent_tab)

    return respondent_tabs


def address_header_slider_changed(address_header_count):
    address_header_dropdowns = []

    for i in range(MAX_ADDRESS_HEADER_COUNT):
        address_header_dropdown = gr.Dropdown(
            visible=i < address_header_count
        )

        address_header_dropdowns.append(address_header_dropdown)

    return address_header_dropdowns


def address_header_dropdown_changed(original_excel_data_frame, address_header):
    if original_excel_data_frame is None or address_header is None:
        return [None] * (MAX_ADDRESS_HEADER_COUNT - 1)

    address_header_index = original_excel_data_frame.columns.get_loc(address_header)

    address_header_dropdowns = []

    for i in range(1, MAX_ADDRESS_HEADER_COUNT):
        address_header_dropdown = gr.Dropdown(
            value=original_excel_data_frame.columns[min(address_header_index + i, len(original_excel_data_frame.columns) - 1)]
        )

        address_header_dropdowns.append(address_header_dropdown)

    return address_header_dropdowns


def clean_button_clicked(original_excel_file_path, original_excel_data_frame, respondent_count, *inputs):
    name_headers = inputs[0:MAX_RESPONDENT_COUNT]
    address_header_counts = inputs[MAX_RESPONDENT_COUNT:2 * MAX_RESPONDENT_COUNT]
    address_headers = inputs[2 * MAX_RESPONDENT_COUNT:]

    gemini_client = genai.Client()

    cleaned_excel_data_frame = pd.DataFrame()

    for i in range(respondent_count):
        name_header = name_headers[i]
        address_header_count = address_header_counts[i]
        address_header_list = address_headers[i * MAX_ADDRESS_HEADER_COUNT:i * MAX_ADDRESS_HEADER_COUNT + address_header_count]

        respondent_strings = []
        respondent_row_indexes = []

        for row_index, row in original_excel_data_frame.iterrows():
            name = str(row.loc[name_header]).strip()

            if name == "None" or len(name) <= 1:
                continue

            addresses = [row.loc[address_header] for address_header in address_header_list]

            respondent_string = str(name)

            for address in addresses:
                address = str(address).strip()

                if address == "None" or len(address) == 0:
                    continue

                respondent_string += ", " + str(address)

            respondent_strings.append(respondent_string)
            respondent_row_indexes.append(row_index)

        respondent_objects = []

        for j in range(0, len(respondent_strings), GEMINI_MAX_RESPONDENT_COUNT):
            gemini_prompt = f"{GEMINI_PROMPT_PREFIX}\n"

            for k in range(j, min(j + GEMINI_MAX_RESPONDENT_COUNT, len(respondent_strings))):
                gemini_prompt += f"\n{respondent_strings[k]}"

            if DEBUG:
                logging.info(f"gemini_prompt: {gemini_prompt}")

            try:
                gemini_response = gemini_client.models.generate_content(
                    model=GEMINI_MODEL_NAME,
                    contents=gemini_prompt,
                    config={
                        "response_mime_type": "application/json",
                        "response_json_schema": RespondentList.model_json_schema()
                    }
                )

                if DEBUG:
                    logging.info(f"gemini_response: {gemini_response.text}")

                respondent_list_object = RespondentList.model_validate_json(gemini_response.text)

                if len(respondent_list_object.respondents) != min(GEMINI_MAX_RESPONDENT_COUNT, len(respondent_strings) - j):
                    raise ValueError("No. of respondents in Gemini response does not match with the no. of respondents in the prompt.")

                for respondent_object in respondent_list_object.respondents:
                    respondent_objects.append(respondent_object)
            except Exception as e:
                logging.error(e)

        for j in range(len(respondent_objects)):
            respondent_row_index = respondent_row_indexes[j]
            respondent_object = respondent_objects[j]

            cleaned_excel_data_frame.loc[respondent_row_index, f"Respondent {i + 1} Name"] = respondent_object.name
            cleaned_excel_data_frame.loc[respondent_row_index, f"Respondent {i + 1} Address Line 1"] = respondent_object.address_line_1
            cleaned_excel_data_frame.loc[respondent_row_index, f"Respondent {i + 1} Address Line 2"] = respondent_object.address_line_2
            cleaned_excel_data_frame.loc[respondent_row_index, f"Respondent {i + 1} Address Line 3"] = respondent_object.address_line_3
            cleaned_excel_data_frame.loc[respondent_row_index, f"Respondent {i + 1} District"] = respondent_object.district
            cleaned_excel_data_frame.loc[respondent_row_index, f"Respondent {i + 1} State"] = respondent_object.state
            cleaned_excel_data_frame.loc[respondent_row_index, f"Respondent {i + 1} PIN Code"] = respondent_object.pin_code

    for i in range(respondent_count):
        name_header = name_headers[i]
        address_header_count = address_header_counts[i]
        address_header_list = address_headers[i * MAX_ADDRESS_HEADER_COUNT:i * MAX_ADDRESS_HEADER_COUNT + address_header_count]

        other_headers = []

        for column_index, column_header in enumerate(original_excel_data_frame.columns):
            if column_header != name_header and column_header not in address_header_list:
                other_headers.append(column_header)

        cleaned_excel_data_frame[other_headers] = original_excel_data_frame[other_headers]

    cleaned_excel_file_path = original_excel_file_path.name.replace(".xlsx", " - Cleaned.xlsx")

    cleaned_excel_data_frame.to_excel(cleaned_excel_file_path, index=False)

    download_button = gr.DownloadButton(
        value=cleaned_excel_file_path
    )

    return [cleaned_excel_data_frame, download_button]


def test_button_clicked():
    original_excel_file_path = "Sample Data 1.xlsx"

    original_excel_data_frame = pd.read_excel(original_excel_file_path)

    respondent_count = 2

    original_excel_column_headers = original_excel_data_frame.columns.tolist()

    name_header_dropdowns = []

    for i in range(MAX_RESPONDENT_COUNT):
        name_header_dropdown = gr.Dropdown(
            choices=original_excel_column_headers
        )

        name_header_dropdowns.append(name_header_dropdown)

    address_header_counts = []

    for i in range(MAX_RESPONDENT_COUNT):
        address_header_counts.append(1)

    address_header_dropdowns = []

    for i in range(MAX_RESPONDENT_COUNT):
        for j in range(MAX_ADDRESS_HEADER_COUNT):
            address_header_dropdown = gr.Dropdown(
                choices=original_excel_column_headers
            )

            address_header_dropdowns.append(address_header_dropdown)

    name_header_dropdowns[0] = gr.Dropdown(
        choices=original_excel_column_headers,
        value="APPLICANT NAME",
    )

    address_header_counts[0] = 9

    address_header_dropdowns[0] = gr.Dropdown(
        choices=original_excel_column_headers,
        value="APPLICANT FATHER NAME ",
    )

    name_header_dropdowns[1] = gr.Dropdown(
        choices=original_excel_column_headers,
        value="CO-APPLICANT NAME",
    )

    address_header_counts[1] = 9

    address_header_dropdowns[10] = gr.Dropdown(
        choices=original_excel_column_headers,
        value="CO APPLICANT FATHER NAME ",
    )

    return [original_excel_file_path, original_excel_data_frame, respondent_count, *name_header_dropdowns, *address_header_counts, *address_header_dropdowns]


with gr.Blocks(title="CNICA Excel Cleaner") as app:
    gr.Markdown("# CNICA Excel Cleaner #")

    gr.Markdown("### Step 1: Upload Excel File ###")

    original_excel_file = gr.File(
        label="Excel file",
        file_types=[".xlsx", ".xls"],
        interactive=True
    )

    original_excel_data_frame = gr.DataFrame(
        label="Excel data",
        headers=[""],
        interactive=False
    )

    gr.Markdown("### Step 2: Select No. of Respondents ###")

    respondent_slider = gr.Slider(
        label="No. of Respondents",
        minimum=1,
        maximum=MAX_RESPONDENT_COUNT,
        step=1,
        interactive=True
    )

    gr.Markdown("### Step 3: Select Respondent's Name and Address Column Headers ###")

    respondent_tabs = []
    name_header_dropdowns = []
    address_header_sliders = []
    address_header_dropdowns = []

    for i in range(MAX_RESPONDENT_COUNT):
        with gr.Tab(
            label=f"Respondent {i + 1}",
            visible=i == 0
        ) as respondent_tab:
            with gr.Row():
                with gr.Column():
                    name_header_dropdown = gr.Dropdown(
                        label="Name column header",
                        interactive=True
                    )

                    name_header_dropdowns.append(name_header_dropdown)

                with gr.Column():
                    address_header_slider = gr.Slider(
                        label="No. of Address column headers",
                        minimum=1,
                        maximum=MAX_ADDRESS_HEADER_COUNT,
                        step=1,
                        interactive=True
                    )

                    address_header_sliders.append(address_header_slider)

                    for j in range(MAX_ADDRESS_HEADER_COUNT):
                        address_header_dropdown = gr.Dropdown(
                            label=f"Address column header {j + 1}",
                            interactive=True,
                            visible=j == 0
                        )

                        address_header_dropdowns.append(address_header_dropdown)

        respondent_tabs.append(respondent_tab)

    gr.Markdown("### Step 4: Clean Excel File ###")

    clean_button = gr.Button(
        value="Clean",
        variant="primary",
        interactive=True
    )

    cleaned_excel_data_frame = gr.DataFrame(
        label="Cleaned Excel data",
        headers=[""],
        interactive=False
    )

    gr.Markdown("### Step 5: Download Excel File ###")

    download_button = gr.DownloadButton(
        value="Download",
        variant="primary",
        interactive=True
    )

    original_excel_file.upload(
        fn=original_excel_file_uploaded,
        inputs=original_excel_file,
        outputs=[original_excel_data_frame, *name_header_dropdowns, *address_header_dropdowns]
    )

    respondent_slider.change(
        fn=respondent_slider_changed,
        inputs=respondent_slider,
        outputs=respondent_tabs
    )

    for i in range(MAX_RESPONDENT_COUNT):
        address_header_slider = address_header_sliders[i]
        address_header_slider.change(
            fn=address_header_slider_changed,
            inputs=address_header_slider,
            outputs=address_header_dropdowns[i * 10:i * 10 + 10]
        )

    for i in range(MAX_RESPONDENT_COUNT):
        address_header_dropdown = address_header_dropdowns[i * 10]
        address_header_dropdown.change(
            fn=address_header_dropdown_changed,
            inputs=[original_excel_data_frame, address_header_dropdown],
            outputs=address_header_dropdowns[i * 10 + 1:i * 10 + 10]
        )

    clean_button.click(
        fn=clean_button_clicked,
        inputs=[original_excel_file, original_excel_data_frame, respondent_slider, *name_header_dropdowns, *address_header_sliders, *address_header_dropdowns],
        outputs=[cleaned_excel_data_frame, download_button]
    )

    if DEBUG:
        test_button = gr.Button(
            value="Test",
            interactive=True
        )

        test_button.click(
            fn=test_button_clicked,
            outputs=[original_excel_file, original_excel_data_frame, respondent_slider, *name_header_dropdowns, *address_header_sliders, *address_header_dropdowns],
        )

app.launch(
    theme=gr.themes.Default(
        primary_hue=gr.themes.colors.blue
    )
)
