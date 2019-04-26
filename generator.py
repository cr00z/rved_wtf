from docxtpl import DocxTemplate
import openpyxl


def get_channel_from_xlsx(xlsx_path, num_of_rows):
    excel_doc = openpyxl.load_workbook(xlsx_path)
    excel_sheet = excel_doc['Sheet1']
    multiple_cells = excel_sheet['A1':'I' + str(num_of_rows)]
    keys = [
        'channel_num',
        'channel_point1',
        'channel_point2',
        'channel_net',
        'channel_type',
        'channel_sa',
        'channel_channel',
        'channel_order',
        'channel_date'
    ]
    for values in multiple_cells:
        yield dict(zip(keys,[cell.value for cell in values]))
        
        
def save_template_as_channel(context):
    docx_name = '{c[channel_num]} {c[channel_point1]}-{c[channel_point2]}.docx' \
        .format(c=context)
    docx_template = DocxTemplate('template.docx')
    docx_template.render(context)
    docx_template.save(docx_name)
        
        
if __name__ == '__main__':
    for context in get_channel_from_xlsx('test.xlsx', 2):
        save_template_as_channel(context)