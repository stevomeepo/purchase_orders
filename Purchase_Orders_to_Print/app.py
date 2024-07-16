import pandas as pd
import os
import re
import fitz
from flask import Flask, render_template, request, send_file
from datetime import date

import tempfile
import ast
import io

app = Flask(__name__)

output_text = ""

@app.route('/')
def index():
    global output_text
    output_text = ""  # Reset output_text
    return render_template('index.html')

@app.route('/highlight_pdf', methods=['POST'])
def highlight_pdf():
    global output_text
    print("Program executing: Purchase_Orders_to_Print")
    # output_text = ""
    excel_file_original = request.files['excel_file']
    pdf_file = request.files['pdf_file']

    excel_file = Receiving_report(excel_file_original)

    with tempfile.NamedTemporaryFile(delete=False) as excel_tempfile:
        excel_path = excel_tempfile.name
        with open(excel_path, 'wb') as file:
            file.write(excel_file.getbuffer())

    with tempfile.NamedTemporaryFile(delete=False) as pdf_tempfile:
        pdf_path = pdf_tempfile.name
        pdf_file.save(pdf_path)

    df = pd.read_excel(excel_path, skiprows=2)
    df = df.rename(columns={"Unnamed: 1": "PO.NO", "Unnamed: 2": "Model", "Unnamed: 3": "QTY"})

    # Obtain the highlight color from the user input
    highlight_color_str = request.form['highlight_color']
    highlight_color = ast.literal_eval(highlight_color_str)
    highlight_color_model = list(highlight_color)

    for i in range(len(df['PO.NO'])):
        po = df.loc[i, 'PO.NO']
        if "TG" in str(po):
            if "TG2" not in str(po):
                string = po.strip()
                parts = re.split(r'\s+', string)
                number = parts[-1]
                number = number.lstrip('0')
                padded_number = number.zfill(10)
                new_po = 'PO' + padded_number
                df.loc[i, 'PO.NO'] = new_po
        if "LPO" not in str(po):
            if 'L' in po:
                if "LD2" not in str(po):
                    string = po.strip()
                    parts = re.split(r'\s+', string)
                    number = parts[-1]
                    number = number.lstrip('0')
                    padded_number = number.zfill(10)
                    new_po = 'LPO' + padded_number
                    df.loc[i, 'PO.NO'] = new_po
        # elif re.match(r'^LPO \d+$', po):
        #     number = re.findall(r'\d+', po)[0]
        #     padded_number = number.zfill(10)
        #     new_po = 'LPO' + padded_number
        #     df.loc[i, 'PO.NO'] = str(new_po)

    class PO:
        def __init__(self, PO_num, Model, QTY):
            self.PO_num = PO_num
            self.Model = Model
            self.QTY = QTY

    PO_list = []
    for index, row in df.iterrows():
        po = PO(row['PO.NO'], row['Model'], row['QTY'])
        PO_list.append(po)

    PO_dict = {}
    for po in PO_list:
        if po.PO_num not in PO_dict:
            PO_dict[po.PO_num] = []
        PO_dict[po.PO_num].append(po)

    model_count_ref = 0
    for po_num in PO_dict:
        # output_text += f"PO_num: {po_num}\n"
        for po in PO_dict[po_num]:
            model_count_ref += 1
            # output_text += f"Model: {po.Model}\n"
            # output_text += f"QTY: {po.QTY}\n"
    # output_text += f"Model_reference: {model_count_ref}\n"

    def highlight_text(pdf_path, PO_dict, output_path, model_count_ref):
        global output_text
        doc = fitz.open(pdf_path)
        pages_to_delete = []
        model_count = 0
        model_be_found = []

        for page in doc:
            model_found = False
            for po_num in PO_dict:
                text_PO = page.search_for(str(po_num))
                if text_PO:
                    for po in PO_dict[po_num]:
                        model = str(po.Model)
                        qty = po.QTY
                        qty = "{:,.0f}".format(qty)
                        qty = str(qty)
                        text_model = page.search_for(model)
                        if text_model:
                            model_found = True
                            for inst in text_model:
                                found_text_model = page.get_textbox(inst).strip()
                                if found_text_model == model:
                                    model_end = inst.x1
                                    highlight = page.add_highlight_annot(inst)
                                    highlight.set_colors(stroke=highlight_color_model)
                                    highlight.update()
                                    # output_text += f"Model found: {found_text_model}\n"
                                    model_be_found.append(found_text_model)
                                    model_count += 1
                                    search_region = fitz.Rect(model_end, inst.y0, page.rect.width, inst.y1)
                                    text_qty = page.search_for(qty, clip=search_region)
                                    if text_qty:
                                        for qty_inst in text_qty:
                                            found_text_qty = page.get_textbox(qty_inst).strip()
                                            qty_start = qty_inst.x1
                                            if found_text_qty == qty:
                                                highlight = page.add_highlight_annot(qty_inst)
                                                highlight.set_colors(stroke=highlight_color_model)
                                                highlight.update()
                                    else:
                                        text_quantity = page.search_for("Quantity")
                                        for quant_inst in text_quantity:
                                            quant_start = quant_inst.x0
                                        qty_x = (model_end + quant_start) / 2
                                        qty_y = inst.y0
                                        qty_width = -8.5
                                        qty_height = inst.y1 - inst.y0 - 2
                                        page.draw_rect(fitz.Rect(qty_x + qty_width, qty_y + 2, qty_x + qty_width + 38, qty_y + qty_height + 2), fill=highlight_color)
                                        page.insert_text(fitz.Point(qty_x + qty_width + 2, qty_y + qty_height), qty, fontsize=10)

            if not model_found:
                pages_to_delete.append(page.number)

        for page_num in reversed(pages_to_delete):
            doc.delete_page(page_num)

        doc.save(output_path)
        doc.close()
        
        miss_found = []
        for i in PO_list:
            if str(i.Model) not in model_be_found:
                miss_found.append(f"{i.PO_num}: {i.Model} QTY: {i.QTY}")

        output_text += "------------------------------------------\n"
        output_text += f"Item_reference_count: {model_count_ref}\n"
        output_text += f"Item_count: {model_count}\n"
        output_text += f"Miss_item: {miss_found}\n"
        output_text += "------------------------------------------\n"
        if len(miss_found) != 0:
            output_text += "Please check the miss item!"
        else:
            output_text += "Status: Done!"


    def highlight_text_taiwan(pdf_path, PO_dict, output_path, model_count_ref):
        global output_text
        doc = fitz.open(pdf_path)
        pages_to_delete = []
        model_count = 0
        model_be_found = []

        for page in doc:
            model_found = False
            for po_num in PO_dict:
                for po in PO_dict[po_num]:
                    model = str(po.Model)
                    qty = po.QTY
                    qty = "{:,.0f}".format(qty)
                    qty = str(qty)
                    text_model = page.search_for(model)
                    if text_model:
                        model_found = True
                        for inst in text_model:
                            found_text_model = page.get_textbox(inst).strip()
                            if found_text_model == model:
                                model_end = inst.x1
                                highlight = page.add_highlight_annot(inst)
                                highlight.set_colors(stroke=highlight_color_model)
                                highlight.update()
                                # output_text += f"Model found: {found_text_model}\n"
                                model_be_found.append(found_text_model)
                                model_count += 1
                                search_region = fitz.Rect(model_end, inst.y0, page.rect.width, inst.y1)
                                text_qty = page.search_for(qty, clip=search_region)
                                if text_qty:
                                    for qty_inst in text_qty:
                                        found_text_qty = page.get_textbox(qty_inst).strip()
                                        qty_start = qty_inst.x1
                                        if found_text_qty == qty:
                                            highlight = page.add_highlight_annot(qty_inst)
                                            highlight.set_colors(stroke=highlight_color_model)
                                            highlight.update()
                                else:
                                    text_quantity = page.search_for("Quantity")
                                    for quant_inst in text_quantity:
                                        quant_start = quant_inst.x0
                                    qty_x = (model_end + quant_start) / 2
                                    qty_y = inst.y0
                                    qty_width = -8.5
                                    qty_height = inst.y1 - inst.y0 - 2
                                    page.draw_rect(fitz.Rect(qty_x + qty_width, qty_y + 2, qty_x + qty_width + 38, qty_y + qty_height + 2), fill=highlight_color)
                                    page.insert_text(fitz.Point(qty_x + qty_width + 2, qty_y + qty_height), qty, fontsize=10)

            if not model_found:
                pages_to_delete.append(page.number)

        for page_num in reversed(pages_to_delete):
            doc.delete_page(page_num)

        doc.save(output_path)
        doc.close()

        miss_found = []
        for i in PO_list:
            if str(i.Model) not in model_be_found:
                miss_found.append(f"{i.PO_num}: {i.Model} QTY: {i.QTY}")


        output_text += "------------------------------------------\n"
        output_text += f"Item_reference_count: {model_count_ref}\n"
        output_text += f"Item_count: {model_count}\n"
        output_text += f"Miss_item: {miss_found}\n"
        output_text += "------------------------------------------\n"
        
        if len(miss_found) != 0:
            output_text += "Please check the miss item!"
        else:
            output_text += "Status: Done!"  



    today = date.today()
    formatted_date = today.strftime('%m%d%Y')

    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_output_file:
        output_path = temp_output_file.name

    file_name = os.path.basename(excel_file_original.filename).lower()
    print("File name", file_name)
    if "taiwan" in file_name:
        highlight_text_taiwan(pdf_path, PO_dict, output_path, model_count_ref)
    else:
        highlight_text(pdf_path, PO_dict, output_path, model_count_ref)

    os.remove(excel_path)
    os.remove(pdf_path)

    return send_file(output_path, as_attachment=True, download_name=f'PO_to_Print_{formatted_date}.pdf')

def Receiving_report(receiving_report):
    global output_text
    file = receiving_report
    file_path = os.path.join(app.root_path, 'uploads', file.filename)
    file.save(file_path)
    
    sheet_names = ['EN PL', 'Eaton PL', 'TOP to HJ Packing slip', 'TOP to PACIFIC Packing slip', 'EN Packing slip', 'TOP PL', 'LIDER PL', 'Table 1', 'LIDER TO PAC Packing slip', 'IN']
    skip_rows = [9, 9, 11, 11, 11, 9, 9, 0, 11, 0]


    def parse_sheets_from_file(file_path, sheet_names, skip_rows):
        parsed_data = {}
        Row_file = pd.ExcelFile(file_path)
        
        for sheet_name, skip_row in zip(sheet_names, skip_rows):
            if sheet_name in Row_file.sheet_names:
                parsed_data[sheet_name] = Row_file.parse(sheet_name=sheet_name, skiprows=skip_row)
        
        return parsed_data


    parsed_data = parse_sheets_from_file(file_path, sheet_names, skip_rows)

    if 'EN PL' in parsed_data:
        EN_PL = parsed_data['EN PL']
        #EN_PL sorting
        EN_PL.rename(columns={'THE CONTENT OF EACH BULK': 'Model', 'PO NO': 'PO.NO'}, inplace=True)
        EN_PL_res = EN_PL[['Model', 'QTY', 'PO.NO']]
        EN_PL_res = EN_PL_res.groupby(['PO.NO', 'Model']).agg(['sum'])
        EN_PL_res = EN_PL_res.reset_index()
    
    if 'Eaton PL' in parsed_data:
        Eaton_PL = parsed_data['Eaton PL']
        #EN_PL sorting
        Eaton_PL.rename(columns={'THE CONTENT OF EACH BULK': 'Model', 'PO NO': 'PO.NO'}, inplace=True)
        Eaton_PL_res = Eaton_PL[['Model', 'QTY', 'PO.NO']]
        Eaton_PL_res = Eaton_PL_res.groupby(['PO.NO', 'Model']).agg(['sum'])
        Eaton_PL_res = Eaton_PL_res.reset_index()

    if 'TOP to HJ Packing slip' in parsed_data:
        TOP_HJ = parsed_data['TOP to HJ Packing slip']
        #TOP_HJ sorting
        TOP_HJ.rename(columns={'PO NO：': 'PO.NO'}, inplace=True)
        TOP_HJ_res = TOP_HJ[['Model', 'QTY', 'PO.NO']]
        TOP_HJ_res = TOP_HJ_res.groupby(['PO.NO', 'Model']).agg(['sum'])
        TOP_HJ_res = TOP_HJ_res.reset_index()

    if 'TOP to PACIFIC Packing slip' in parsed_data:
        TOP_PAC = parsed_data['TOP to PACIFIC Packing slip']
        #TOP_PAC sorting
        TOP_PAC_res = TOP_PAC[['Model', 'QTY', 'PO.NO']]
        TOP_PAC_res = TOP_PAC_res.groupby(['PO.NO', 'Model']).agg(['sum'])
        TOP_PAC_res = TOP_PAC_res.reset_index()

    if 'EN Packing slip' in parsed_data:
        EN_PACK = parsed_data['EN Packing slip']
        #EN_PAC sorting
        EN_PACK_res = EN_PACK[['Model', 'QTY', 'PO.NO']]
        EN_PACK_res = EN_PACK_res.groupby(['PO.NO', 'Model']).agg(['sum'])
        EN_PACK_res = EN_PACK_res.reset_index()

    if 'TOP PL' in parsed_data:
        TOP_PL = parsed_data['TOP PL']
        #TOP_PL sorting
        TOP_PL.rename(columns={'THE CONTENT OF EACH BULK': 'Model', 'PO NO': 'PO.NO'}, inplace=True)
        TOP_PL_res = TOP_PL[['Model', 'QTY', 'PO.NO']]
        TOP_PL_res = TOP_PL_res.groupby(['PO.NO', 'Model']).agg(['sum'])
        TOP_PL_res = TOP_PL_res.reset_index()

    if 'LIDER PL' in parsed_data:
        LIDER_PL = parsed_data['LIDER PL']
        #LIDER_PL sorting
        LIDER_PL.rename(columns={'THE CONTENT OF EACH BULK': 'Model', 'PO NO': 'PO.NO'}, inplace=True)
        LIDER_PL_res = LIDER_PL[['Model', 'QTY', 'PO.NO']]
        LIDER_PL_res = LIDER_PL_res.groupby(['PO.NO', 'Model']).agg(['sum'])
        LIDER_PL_res = LIDER_PL_res.reset_index()

    if 'Table 1' in parsed_data:
        Table_1 = parsed_data['Table 1']
        #Table_1 sorting
        PO = Table_1.iloc[3][7]
        Table_1 = Table_1.iloc[6:].copy()
        Table_1.rename(columns={'Unnamed: 1': 'Model', 'Unnamed: 2': 'QTY'}, inplace=True)
        Table_1_res = Table_1[['Model', 'QTY']].dropna()
        Table_1_res = Table_1_res.groupby(['Model']).agg(['sum'])
        Table_1_res = Table_1_res.reset_index()
        Table_1_res = Table_1_res[Table_1_res['Model'] != 'Amount:']
        Table_1_res.insert(0, 'PO.NO', PO)

    if 'LIDER TO PAC Packing slip' in parsed_data:
        LID_PAC = parsed_data['LIDER TO PAC Packing slip']
        #LID_PAC sorting
        LID_PAC_res = LID_PAC[['Model', 'QTY', 'PO.NO']]
        LID_PAC_res = LID_PAC_res.groupby(['PO.NO', 'Model']).agg(['sum'])
        LID_PAC_res = LID_PAC_res.reset_index()
    
    if 'IN' in parsed_data:
        output_text += "Please highlight the sheet name \"IN\" manually!\n"

    dataframes = []

    if 'EN_PL_res' in locals():
        dataframes.append(EN_PL_res)
    if 'Eaton_PL_res' in locals():
        dataframes.append(Eaton_PL_res)
    if 'EN_PACK_res' in locals():
        dataframes.append(EN_PACK_res)
    if 'TOP_HJ_res' in locals():
        dataframes.append(TOP_HJ_res)
    if 'TOP_PAC_res' in locals():
        dataframes.append(TOP_PAC_res)
    if 'TOP_PL_res' in locals():
        dataframes.append(TOP_PL_res)
    if 'LIDER_PL_res' in locals():
        dataframes.append(LIDER_PL_res)
    if 'Table_1_res' in locals():
        dataframes.append(Table_1_res)
    if 'LID_PAC_res' in locals():
        dataframes.append(LID_PAC_res)

    concatenated_df = pd.concat(dataframes, axis=0)
    # concatenated_df = concatenated_df.groupby(['PO.NO', 'Model']).agg('sum')

    # Rest of your code for processing and aggregating the data

    output_file_name = 'Receiving_report'
    output_file_path = os.path.join(app.root_path, 'uploads', f'{output_file_name}_{file.filename}')
    # concatenated_df.to_excel(output_file_path)

    # Save the processed file to a BytesIO object
    output_file = io.BytesIO()
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        concatenated_df.to_excel(writer)
    
    # Reset the file pointer of the BytesIO object
    output_file.seek(0)



    # Send the temporary file as a response for the user to download
    return output_file



@app.route('/get_output_text')
def get_output_text():
    global output_text
    return output_text


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001)