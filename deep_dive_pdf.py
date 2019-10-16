# -*- coding: utf-8 -*-
"""
Created on Thu May 23 17:12:18 2019

@author: pbonnin
"""

def auto_filename():
    import datetime
    today = datetime.date.today()
    first = today.replace(day=1)
    lastMonth = first - datetime.timedelta(days=1)
    current_month = lastMonth.strftime("%B %Y")
    return(current_month)
    
def get_filename():
    # a list of months to check againt
    months = ['January','February','March','April','May','June','July','August','September','October','November','December']
    years = [2018, 2019]
    check = [x + ' ' + str(y) for x in months for y in years]
    
    while True:
        title = input('\n'+'Enter alternate month and year for the report (e.g. June 2019): ')
        if title.title() in check:
            return(title.title())
            break
        elif title == 'quit':
            break
        else:
            print('\n','Invalid entry',sep='')

def output_pdf(infile,outfile,excl_list):
    from PyPDF2 import PdfFileWriter, PdfFileReader
    temp_output = PdfFileWriter()
    infile = PdfFileReader(infile, 'rb')
    for i in excl_list:
        p = infile.getPage(i)
        temp_output.addPage(p)

    with open(outfile, 'wb') as f:
        temp_output.write(f)


def clean_cell(cell_text):
    temp = str(cell_text).replace('text:','').replace("'",'')
    temp = temp.split(',')
    temp = [int(i) for i in temp]
    return(temp)


def confirm_filename():
    while True:
        input_statement = 'Please confirm '+auto_filename()+' is the correct report (Y/N):'
        get = input(input_statement)
        if get.upper() == 'Y':
            get = auto_filename()
            break
        elif get.upper() == 'N':
            while True:
                get = get_filename()
                break
            break
        else:
            print('Please enter only Y or N')
            continue
    return(get.upper())



def main(current_month):
    from PyPDF2 import PdfFileWriter, PdfFileReader
    import xlrd
    import time
    
    current_month = current_month.title()
    
    start_time = time.time()

    workbook = xlrd.open_workbook('R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/Excel/Cinemax Monthly Report - '+current_month+'.xlsx', on_demand = True)
    
    worksheet = workbook.sheet_by_name('REF')
    
    # Imports the index lists
    pages_to_keep_fr = clean_cell(worksheet.cell(15, 101))
    pages_to_keep_ddd = clean_cell(worksheet.cell(24, 101))
    pages_to_keep_pan = clean_cell(worksheet.cell(33, 101))
    pages_to_keep_mex = clean_cell(worksheet.cell(37, 101))
    pages_to_keep_bra = clean_cell(worksheet.cell(41, 101))
    pages_to_keep_reps = clean_cell(worksheet.cell(45, 101))
    new_order = clean_cell(worksheet.cell(63, 101))


    #these dont change
    
    pages_to_keep_dd = [0, 1, 8, 9, 16, 17, 24, 25, 32, 33, 40, 41, 48, 49, 56, 57]
    
    pages_to_keep_gn = [0, 1, 2, 3, 4, 5, 6]
    
    # the deeper dive
    dd_order = ['PAN','MEX','BR','AR','CO', 'CH', 'PE', 'CA', 'REPS']
    
    dd_pan = list(range(0,7))
    dd_mex = list(range(8,15))
    dd_bra = list(range(16,23))
    dd_ar = list(range(24,31))
    dd_co = list(range(32,39))
    dd_ch = list(range(40,47))
    dd_pe = list(range(48,55))
    dd_ca = list(range(56,63))
    dd_re = list(range(24,63))
    
    dd_indexes =[dd_pan,dd_mex,dd_bra,dd_ar,dd_co,dd_ch,dd_pe,dd_ca,dd_re]
    
    # the good news
    gn_order = ['PAN','MEX','BR','REPS']
    gn_indexes = [[0],[1],[6],[2,3,4,5]]


    full_report = 'R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/PDF/Raw/Cinemax Monthly Report - '+current_month+'.pdf'
    
    deeper_dive = 'R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/Deeper Dive PDF/Deeper Dive Rankers - '+current_month+' (ALL).pdf'
    
    good_news = 'R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/PDF - Good News!/Cinemax Monthly Round-Up - '+current_month+'.pdf'
    
    output_file = 'R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/TMV Deep Dive/Cinemax Deep Dive Meeting Compilation - '+current_month+'.pdf'
        
    infile = PdfFileReader(full_report, 'rb')
    infile2 = PdfFileReader(deeper_dive, 'rb')
    infile3 = PdfFileReader(good_news, 'rb')
    
    output = PdfFileWriter()
    
    for i in pages_to_keep_fr:
        p = infile.getPage(i)
        output.addPage(p)
        
    for i in pages_to_keep_dd:
        p = infile2.getPage(i)
        output.addPage(p)
        
    for i in pages_to_keep_gn:
        p = infile3.getPage(i)
        output.addPage(p)
    
    with open(output_file, 'wb') as f:
        output.write(f)
        
    output2 = PdfFileWriter()
    
    infile4 = PdfFileReader(output_file, 'rb')
    
    for i in new_order:
        p = infile4.getPage(i)
        output2.addPage(p)
    
    with open(output_file, 'wb') as f:
        output2.write(f)

    # output PDFs
    output_ddd = 'R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/PDF/Cinemax Monthly Report - '+current_month+'.pdf'
    output_pdf(full_report,output_ddd,pages_to_keep_ddd)
    
    output_pan = 'R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/PDF/Cinemax Monthly Report - '+current_month+' (PAN).pdf'
    output_pdf(full_report,output_pan,pages_to_keep_pan)
    
    output_mex = 'R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/PDF/Cinemax Monthly Report - '+current_month+' (MEX).pdf'
    output_pdf(full_report,output_mex,pages_to_keep_mex)
    
    output_bra = 'R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/PDF/Cinemax Monthly Report - '+current_month+' (BR).pdf'
    output_pdf(full_report,output_bra,pages_to_keep_bra)
    
    output_reps = 'R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/PDF/Cinemax Monthly Report - '+current_month+' (REPS).pdf'
    output_pdf(full_report,output_reps,pages_to_keep_reps)

    # Deeper Dive and Good News individual files
    
    for region, array in zip(dd_order,dd_indexes):
        filename = 'R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/Deeper Dive PDF/Deeper Dive Rankers - '+current_month+' ('+region+').pdf'
        output_pdf(deeper_dive,filename,array)
        
    
    for region, array in zip(gn_order,gn_indexes):
        filename = 'R:/Cinemax Research/Ad Sales/Monthy Deep-Dive Reports/PDF - Good News!/Cinemax Monthly Round-Up - '+current_month+' ('+region+').pdf'
        output_pdf(good_news,filename,array)

    print('Process complete:',str(round(time.time() - start_time,2)),'seconds')

if __name__== "__main__":
  main(confirm_filename())