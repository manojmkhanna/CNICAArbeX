import logging

import gradio as gr
import pandas as pd

from google import genai
from pydantic import BaseModel


RESP_MAX_COUNT = 10
ADDR_MAX_COUNT = 10
GEMINI_MODEL_NAME = "gemini-3-flash-preview"
GEMINI_PROMPT_PREFIX = "Split the following row of addresses into columns such as Recipient Name/Entity Name, Address Line 1/Care of Name, Address Line 2, Address Line 3, District, State and PIN Code. " \
    "For Name and Care of Name add salutations like Mr., Mrs., Ms. or M/s. if missing. " \
    "For Care of Name add prefixes like s/o, d/o, f/o, m/o, h/o, w/o or c/o if missing. " \
    "Correct spelling mistakes and punctuations in an address if necessary. " \
    "Correct an incomplete address if necessary. " \
    "Remove redundancy in an address if necessary. " \
    "Convert everything to Proper case."
GEMINI_MAX_RESP_COUNT = 50


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


def excel_file_changed(excel_path):
    if excel_path is None:
        return [None] * (1 + RESP_MAX_COUNT + (RESP_MAX_COUNT * ADDR_MAX_COUNT))

    excel_data = pd.read_excel(excel_path)

    excel_column_headers = excel_data.columns.tolist()

    name_dropdowns = []

    for i in range(RESP_MAX_COUNT):
        name_dropdown = gr.Dropdown(
            choices=excel_column_headers,
            value=excel_column_headers[0]
        )

        name_dropdowns.append(name_dropdown)

    addr_dropdowns = []

    for i in range(RESP_MAX_COUNT):
        for j in range(ADDR_MAX_COUNT):
            addr_dropdown = gr.Dropdown(
                choices=excel_column_headers,
                value=excel_column_headers[0]
            )

            addr_dropdowns.append(addr_dropdown)

    return [excel_data] + name_dropdowns + addr_dropdowns


def resp_slider_changed(resp_count):
    resp_tabs = []

    for i in range(RESP_MAX_COUNT):
        resp_tab = gr.Tab(
            visible=i < resp_count
        )

        resp_tabs.append(resp_tab)

    return resp_tabs


def addr_slider_changed(addr_count):
    addr_dropdowns = []

    for i in range(ADDR_MAX_COUNT):
        addr_dropdown = gr.Dropdown(
            visible=i < addr_count
        )

        addr_dropdowns.append(addr_dropdown)

    return addr_dropdowns


def addr_dropdown_changed(excel_data, addr_column_header):
    if excel_data is None or addr_column_header is None:
        return [None] * (ADDR_MAX_COUNT - 1)

    addr_column_index = excel_data.columns.get_loc(addr_column_header)

    addr_dropdowns = []

    for i in range(1, ADDR_MAX_COUNT):
        addr_dropdown = gr.Dropdown(
            value=excel_data.columns[min(addr_column_index + i, len(excel_data.columns) - 1)]
        )

        addr_dropdowns.append(addr_dropdown)

    return addr_dropdowns


def clean_button_clicked(excel_data, resp_count, *inputs):
    name_column_headers = inputs[0:RESP_MAX_COUNT]
    addr_counts = inputs[RESP_MAX_COUNT:2 * RESP_MAX_COUNT]
    addr_column_headers = inputs[2 * RESP_MAX_COUNT:]

    gemini_client = genai.Client()

    cleaned_excel_data = pd.DataFrame()

    for i in range(resp_count):
        resp_name_column_header = name_column_headers[i]
        resp_name_column_index = excel_data.columns.get_loc(resp_name_column_header)

        resp_addr_count = addr_counts[i]
        resp_addr_column_headers = addr_column_headers[i * ADDR_MAX_COUNT:i * ADDR_MAX_COUNT + resp_addr_count]
        resp_addr_column_indexes = [excel_data.columns.get_loc(resp_addr_column_header) for resp_addr_column_header in resp_addr_column_headers]

        resps = []

        for row_index, row in excel_data.iterrows():
            resp_name = row.iloc[resp_name_column_index]

            if resp_name is None or len(str(resp_name).strip()) <= 3:
                continue

            resp_addrs = [row.iloc[column_index] for column_index in resp_addr_column_indexes]

            resp = str(resp_name) if resp_name is not None else ""
            for resp_addr in resp_addrs:
                resp += ", " + str(resp_addr) if resp_addr is not None else ""

            resps.append((row_index, resp))

        for j in range(0, len(resps), GEMINI_MAX_RESP_COUNT):
            from_index = j
            to_index = min(j + GEMINI_MAX_RESP_COUNT, len(resps))

            gemini_prompt = f"{GEMINI_PROMPT_PREFIX}\n"
            for k in range(from_index, to_index):
                gemini_prompt += f"\n{resps[k][1]}"

            logging.info(f"gemini_prompt: {gemini_prompt}")

            gemini_response = gemini_client.models.generate_content(
                model=GEMINI_MODEL_NAME,
                contents=gemini_prompt,
                config={
                    "response_mime_type": "application/json",
                    "response_json_schema": RespondentList.model_json_schema()
                }
            )

            gemini_answer = RespondentList.model_validate_json(gemini_response.text)

            logging.info(f"gemini_answer: {gemini_answer}")

            for k in range(from_index, to_index):
                row_index = resps[k][0]
                gemini_resp = gemini_answer.respondents[k - from_index]

                cleaned_excel_data.loc[row_index, f"Respondent {i + 1} Name"] = gemini_resp.name
                cleaned_excel_data.loc[row_index, f"Respondent {i + 1} Address Line 1"] = gemini_resp.address_line_1
                cleaned_excel_data.loc[row_index, f"Respondent {i + 1} Address Line 2"] = gemini_resp.address_line_2
                cleaned_excel_data.loc[row_index, f"Respondent {i + 1} Address Line 3"] = gemini_resp.address_line_3
                cleaned_excel_data.loc[row_index, f"Respondent {i + 1} District"] = gemini_resp.district
                cleaned_excel_data.loc[row_index, f"Respondent {i + 1} State"] = gemini_resp.state
                cleaned_excel_data.loc[row_index, f"Respondent {i + 1} PIN Code"] = gemini_resp.pin_code

    for i in range(resp_count):
        resp_name_column_header = name_column_headers[i]
        resp_name_column_index = excel_data.columns.get_loc(resp_name_column_header)

        resp_addr_count = addr_counts[i]
        resp_addr_column_headers = addr_column_headers[i * ADDR_MAX_COUNT:i * ADDR_MAX_COUNT + resp_addr_count]
        resp_addr_column_indexes = [excel_data.columns.get_loc(resp_addr_column_header) for resp_addr_column_header in resp_addr_column_headers]

        ignore_indexes = []
        ignore_indexes.append(resp_name_column_index)
        for resp_addr_column_index in resp_addr_column_indexes:
            ignore_indexes.append(resp_addr_column_index)

        for col_index, column in enumerate(excel_data.columns):
            if col_index not in ignore_indexes:
                cleaned_excel_data[column] = excel_data[column]

    return cleaned_excel_data


with gr.Blocks(title="CNICA Excel Cleaner") as app:
    gr.Markdown("# CNICA Excel Cleaner #")

    gr.Markdown("### Step 1: Upload Excel File ###")

    excel_file = gr.File(
        label="Excel file",
        file_types=[".xlsx", ".xls"],
        interactive=True
    )

    excel_data_frame = gr.DataFrame(
        label="Excel data",
        headers=[""],
        interactive=False
    )

    gr.Markdown("### Step 2: Select No. of Respondents ###")

    resp_slider = gr.Slider(
        label="No. of Respondents",
        minimum=1,
        maximum=RESP_MAX_COUNT,
        step=1,
        interactive=True
    )

    gr.Markdown("### Step 3: Select Respondent's Name and Address Column Headers ###")

    resp_tabs = []
    name_dropdowns = []
    addr_sliders = []
    addr_dropdowns = []

    for i in range(RESP_MAX_COUNT):
        with gr.Tab(
            label=f"Respondent {i + 1}",
            visible=i == 0
        ) as resp_tab:
            with gr.Row():
                with gr.Column():
                    name_dropdown = gr.Dropdown(
                        label="Name column header",
                        interactive=True
                    )

                    name_dropdowns.append(name_dropdown)

                with gr.Column():
                    addr_slider = gr.Slider(
                        label="No. of Address column headers",
                        minimum=1,
                        maximum=ADDR_MAX_COUNT,
                        step=1,
                        interactive=True
                    )

                    addr_sliders.append(addr_slider)

                    for j in range(ADDR_MAX_COUNT):
                        addr_dropdown = gr.Dropdown(
                            label=f"Address column header {j + 1}",
                            interactive=True,
                            visible=j == 0
                        )

                        addr_dropdowns.append(addr_dropdown)

        resp_tabs.append(resp_tab)

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

    excel_file.change(
        fn=excel_file_changed,
        inputs=excel_file,
        outputs=[excel_data_frame, *name_dropdowns, *addr_dropdowns]
    )

    resp_slider.change(
        fn=resp_slider_changed,
        inputs=resp_slider,
        outputs=resp_tabs
    )

    for i in range(RESP_MAX_COUNT):
        addr_slider = addr_sliders[i]
        addr_slider.change(
            fn=addr_slider_changed,
            inputs=addr_slider,
            outputs=addr_dropdowns[i * 10:i * 10 + 10]
        )

    for i in range(RESP_MAX_COUNT):
        addr_dropdown = addr_dropdowns[i * 10]
        addr_dropdown.change(
            fn=addr_dropdown_changed,
            inputs=[excel_data_frame, addr_dropdown],
            outputs=addr_dropdowns[i * 10 + 1:i * 10 + 10]
        )

    clean_button.click(
        fn=clean_button_clicked,
        inputs=[excel_data_frame, resp_slider, *name_dropdowns, *addr_sliders, *addr_dropdowns],
        outputs=cleaned_excel_data_frame
    )

app.launch()
