import os, sys
from typing import ItemsView  # Standard Python Libraries
import xlwings as xw  # pip install xlwings
from docxtpl import DocxTemplate, RichText  # pip install docxtpl
import pandas as pd  # pip install pandas
import matplotlib.pyplot as plt  # pip install matplotlib
import win32com.client as win32  # pip install pywin32

# -- Documentation:
# python-docx-template: https://docxtpl.readthedocs.io/en/latest/

# Change path to current working directory
os.chdir(sys.path[0])


def convert_to_pdf(doc):
    # Convert given word document to pdf
    word = win32.DispatchEx("Word.Application")
    new_name = doc.replace(".docx", r".pdf")
    worddoc = word.Documents.Open(doc)
    worddoc.SaveAs(new_name, FileFormat=17)
    worddoc.Close()
    return None


def main():
    #wb = xw.Book.caller()
    wb = xw.Book('DCDRaptor.xlsm')

    #set the template DCD file
    doc_template = DocxTemplate('T_XXXX_0001_VP0x DCD Template_SW_DesignChangeDocument.docx')

    # excel read the main part dictionary
    sht_main = wb.sheets['MAIN']
    #main_df = sht_main.range('A1').options(pd.DataFrame, index=False, expand='table').value
    dict_main = sht_main.range('A2').options(dict, expand='table', numbers=int).value
    print(dict_main)

    # set output file name
    file_out_dcd = dict_main['PH_RO_NUMBER'] + '_' + dict_main['PH_RO_PROJECT_LEAD'] + '_' + 'DCD' + '_' + dict_main['PH_RO_NAME'] + '.docx'

    

    # excel read the design part dataframe
    sht_design = wb.sheets['DESIGN']
    design_df = sht_design.range('A1').options(pd.DataFrame, index=False, expand='table').value
    design_df['ph_ro_number'] = dict_main['PH_RO_NUMBER']
    design_df['ph_ro_project_lead'] = dict_main['PH_RO_PROJECT_LEAD']
    design_df['ph_ro_name'] = dict_main['PH_RO_NAME']
    design_df['ph_ro_sw_base'] = dict_main['PH_RO_SW_BASE']
    design_df['ph_ro_date_ceation'] = dict_main['PH_RO_DATE_CREATION']
    design_df['ph_ro_affected_comp'] = dict_main['PH_RO_AFFECTED_COMP']
    design_df['ph_ro_responsible_person'] = dict_main['PH_RO_RESPOSIBLE_NAME']
    design_df['ph_ro_fileout_dcd'] = file_out_dcd
    design_df['ph_ro_description'] = dict_main['PH_RO_DESCRIPTION']
    


    print(design_df)

    #convert into dictionary
    context = {'items' : design_df.to_dict('records')}

    #print(context)


    doc_template.render(context)
    doc_template.save(file_out_dcd)  


    # -- Render & Save MAIN DCD Word Document
    #val_ro_project = 'VP02'
    #output_filename_dcd = val_ro_number + '_' + val_ro_project + '_' + 'DCD' + '_' + val_ro_name + '.docx'
    #sht_main = wb.sheets['MAIN']
    #context_main_context = sht_main.range('A2').options(dict, expand='table', numbers=int).value
    #doc_dcd_template.render(context_main_context)
    #doc_dcd_template.save(output_filename_dcd)
    
    # -- Render & Save Child 1 DCD Word Document
    #val_ro_project = 'VP00'
    #output_filename_dcd = val_ro_number + '_' + val_ro_project + '_' + 'DCD' + '_' + val_ro_name + '.docx'
    #sht_main = wb.sheets['MAIN']
    #sht_main.range('B3').value = val_ro_project
    #context_child1_context = sht_main.range('A2').options(dict, expand='table', numbers=int).value
    #doc_dcd_template.render(context_child1_context)
    #doc_dcd_template.save(output_filename_dcd)
    
    # -- Render & Save child 2 DCD Word Document
    #val_ro_project = 'VP01'
    #output_filename_dcd = val_ro_number + '_' + val_ro_project + '_' + 'DCD' + '_' + val_ro_name + '.docx'
    #sht_main = wb.sheets['MAIN']
    #sht_main.range('B3').value = val_ro_project
    #context_child2_context = sht_main.range('A2').options(dict, expand='table', numbers=int).value
    #doc_dcd_template.render(context_child2_context)
    #doc_dcd_template.save(output_filename_dcd)
 
    

    # -- Convert to PDF [OPTIONAL]
    #path_to_word_document = os.path.join(os.getcwd(), output_name)
    #convert_to_pdf(path_to_word_document)

    # -- Show Message Box [OPTIONAL]
    #show_msgbox = wb.macro('Module1.ShowMsgBox')
    #show_msgbox('DONE!')


    #restore
    # sht_main.range('B2').value = val_ro_number_def
    # sht_main.range('B3').value = val_ro_project_def
    # sht_main.range('B4').value = val_ro_name_def


if __name__ == '__main__':
    xw.Book('DCDRaptor.xlsm').set_mock_caller()
    main()
