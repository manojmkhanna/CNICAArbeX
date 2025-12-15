import gradio as gr
import pandas as pd


RESP_MAX_COUNT = 20
ADDR_MAX_COUNT = 10


excel_data_frame = None
excel_column_headers = None


def excel_file_changed(excel_path):
    if excel_path is None:
        return [None] * (1 + RESP_MAX_COUNT + (RESP_MAX_COUNT * ADDR_MAX_COUNT))

    outputs = []

    global excel_data_frame
    excel_data_frame = pd.read_excel(excel_path)

    outputs.append(excel_data_frame)

    global excel_column_headers
    excel_column_headers = excel_data_frame.columns.tolist()

    for i in range(RESP_MAX_COUNT):
        name_dropdown = gr.Dropdown(
            choices=excel_column_headers,
            value=excel_column_headers[0]
        )

        outputs.append(name_dropdown)

    for i in range(RESP_MAX_COUNT):
        for j in range(ADDR_MAX_COUNT):
            addr_dropdown = gr.Dropdown(
                choices=excel_column_headers,
                value=excel_column_headers[0]
            )

            outputs.append(addr_dropdown)

    return outputs


def resp_slider_changed(resp_count):
    resp_tabs = []

    for i in range(RESP_MAX_COUNT):
        resp_tabs.append(gr.Tab(
            visible=i < resp_count
        ))

    return resp_tabs


def addr_slider_changed(addr_count):
    addr_dropdowns = []

    for i in range(ADDR_MAX_COUNT):
        addr_dropdowns.append(gr.Dropdown(
            visible=i < addr_count
        ))

    return addr_dropdowns


def addr_dropdown_changed(addr_column_header):
    if addr_column_header is None:
        return [None] * (ADDR_MAX_COUNT - 1)

    addr_column_index = excel_column_headers.index(addr_column_header)

    addr_dropdowns = []

    for i in range(1, ADDR_MAX_COUNT):
        addr_dropdowns.append(gr.Dropdown(
            value=excel_column_headers[addr_column_index + i]
        ))

    return addr_dropdowns


with gr.Blocks(title="CNICA Excel Cleaner") as app:
    gr.Markdown("# CNICA Excel Cleaner #")

    gr.Markdown("### Step 1: Select Excel File ###")

    excel_file = gr.File(
        label="Excel file",
        file_types=[".xlsx", ".xls"]
    )

    excel_data_frame = gr.DataFrame(
        label="Excel data",
        headers=[""]
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
    addr_dropdowns_dict = {}

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

                    addr_dropdowns = []

                    for j in range(ADDR_MAX_COUNT):
                        addr_dropdown = gr.Dropdown(
                            label=f"Address column header {j + 1}",
                            interactive=True,
                            visible=j == 0
                        )

                        addr_dropdowns.append(addr_dropdown)

                    addr_dropdowns_dict[i] = addr_dropdowns

                    addr_slider.change(
                        fn=addr_slider_changed,
                        inputs=addr_slider,
                        outputs=addr_dropdowns
                    )

                    addr_dropdowns[0].change(
                        fn=addr_dropdown_changed,
                        inputs=addr_dropdowns[0],
                        outputs=addr_dropdowns[1:]
                    )

        resp_tabs.append(resp_tab)

    resp_slider.change(
        fn=resp_slider_changed,
        inputs=resp_slider,
        outputs=resp_tabs
    )

    excel_file_changed_outputs = []
    excel_file_changed_outputs.append(excel_data_frame)
    for i in range(RESP_MAX_COUNT):
        excel_file_changed_outputs.append(name_dropdowns[i])
    for i in range(RESP_MAX_COUNT):
        for j in range(ADDR_MAX_COUNT):
            excel_file_changed_outputs.append(addr_dropdowns_dict[i][j])

    excel_file.change(
        fn=excel_file_changed,
        inputs=excel_file,
        outputs=excel_file_changed_outputs
    )

app.launch()
