import gradio as gr
import pandas as pd


RESP_MAX_COUNT = 20
ADDR_MAX_COUNT = 10


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

    excel_file.change(
        fn=excel_file_changed,
        inputs=excel_file,
        outputs=[excel_data_frame] + name_dropdowns + addr_dropdowns
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

app.launch()
