import gradio as gr
import pandas as pd


def excel_file_changed(excel_file):
    if excel_file is None:
        return None

    return pd.read_excel(excel_file)


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

    excel_file.change(
        fn=excel_file_changed,
        inputs=excel_file,
        outputs=excel_data_frame
    )

app.launch()
