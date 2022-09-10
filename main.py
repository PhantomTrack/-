# конвертация
import os
import sys

import comtypes.client


def ppt_2_pdf(input_ppt_file, output_pdf_file, format_type=32):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if output_pdf_file[-3:] != 'pdf':
        output_pdf_file = output_pdf_file + ".pdf"

    ppt_file = powerpoint.Presentations.Open(input_ppt_file)
    ppt_file.SaveAs(output_pdf_file, format_type)
    ppt_file.Close()
    powerpoint.Quit()


print('Конвертация...')
ppt_2_pdf('C:/Users/Maxim/PycharmProjects/PhantomID/test.pptx',
          f'C:/Users/Maxim/PycharmProjects/PhantomID/test-pdf')
# удаление остаточных файлов
os.remove('C:/Users/Maxim/PycharmProjects/PhantomID/test.pptx')
