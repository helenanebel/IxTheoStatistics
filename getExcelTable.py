import os
import sys
import time
import xlsxwriter
import csv
import urllib.request
from bs4 import BeautifulSoup
import re

path_to_directory = '//sn00.zdv.uni-tuebingen.de/ZE020110/FID-Projekte/Statistik/Statistik_automatisiert/'


def convert_excel(file_name):
    # progressbar = ttk.Progressbar(window, orient="horizontal", length=300, mode="determinate")
    # progressbar.grid(column=0, row=2, columnspan=2)
    # window.update()
    # progressbar["value"] = 0
    # progressbar["maximum"] = 100
    # currentValue = 0
    # Listen liegen hier: W:\FID-Projekte\Team Retro-Scan\ImageWare\1 Vorarbeit MyBib eDoc\Prüfung PPN-Listen
    csv_file = open(path_to_directory + '/' + file_name, 'r', encoding='utf-8')
    csv_reader = csv.reader(csv_file, delimiter=';', quotechar='"')
    print(csv_reader)
    rows = [row for row in csv_reader]
    items_of_series_data = []
    selected_data = []
    data_of_items_with_articles = []
    request_not_successfull = []
    row_nr = 0
    header = []
    header_for_items_with_articles = []
    total_rows = len(rows)
    units = int(total_rows/100)
    counter = 0
    total = 0
    for row in rows:
        '''if stop_processing:
            break'''
        row_nr += 1
        if row_nr == 1:
            header = [{'header': column} for column in row]
            header_for_items_with_articles = [{'header': column} for column in row]
            header_for_items_with_articles.insert(1, {'header': 'Anzahl vorhandener Artikel'})
        if row_nr < 2:
            continue
        '''if counter % units == 0:
            if stop_processing:
                window.quit()
                break
            currentValue = currentValue + 1
            progressbar.after(500, progress(progressbar, currentValue))
            progressbar.update()
            counter = 0'''
        counter += 1
        retrieve_sign = row[1]
        is_item_of_series = False
        retrieve_signs_exk = ["ixau", "ixmo", "ixsb", "ixze", "ixzg", "ixzs", "ixzw", "ixzx", "ixrk", "rwzw", "rwzx",
                              "rwrk", "ixzo", "zota", "imwa", "ixbt", "rwex", "bril", "gruy",
                              "knix", "kn28", "mszo", "mszk", "msmi", "redo", "bsbo"]
        retrieve_signs_cod = ["DTH5", "mteo", "mtex", "BIIN", "KALD", "GIRA", "DAKR", "AUGU", "MIKA"]
        retrieve_signs_sfk = ["1", "0", "6,22", "1 or 0", "1 or 0 or 6,22", "1 or 0 or 6,22 or mteo not mtex"]
        skip = ["IxTheo", "RelBib", "IxBib", "KALDI/DaKaR", "Seit August 2020 >2.300:"]
        if retrieve_sign:
            total += 1
            try:
                if retrieve_sign in retrieve_signs_exk:
                    filedata = urllib.request.urlopen('http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.exk%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml')
                elif retrieve_sign in retrieve_signs_cod:
                    filedata = urllib.request.urlopen('http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.cod%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml')
                elif retrieve_sign in retrieve_signs_sfk:
                    if 'or' not in retrieve_sign:
                        filedata = urllib.request.urlopen(
                        'http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.sfk%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml')
                    elif retrieve_sign == "1 or 0 or 6,22 or mteo not mtex":
                        filedata = urllib.request.urlopen("http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.sfk%3D0+or+pica.sfk%3D1+or+pica.sfk%3D6,22+or+pica.cod%3Dmteo+not+pica.cod%3Dmtex&maximumRecords=10&recordSchema=picaxml")
                    else:
                        retrieve_sign_list = retrieve_sign.split(' or ')
                        retrieve_sign = "+or+pica.sfk%3D".join(retrieve_sign_list)
                        filedata = urllib.request.urlopen(
                            'http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.sfk%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml')
                        print('http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.sfk%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml')
                elif retrieve_sign == "IxTheo":
                    filedata = urllib.request.urlopen("https://www.ixtheo.de/Search/Results?lookfor=&type=AllFields&botprotect=")
                    data = filedata.read()
                    xml_soup = BeautifulSoup(data, "lxml")
                    number_of_records = re.findall(r'results of ([\d,]*) for search', xml_soup.find('div', class_='search-stats').text)
                    if number_of_records:
                        number_of_records = number_of_records[0].replace(',', '')
                        print(number_of_records)
                    continue
                elif retrieve_sign == "RelBib":
                    # "https://www.relbib.de/Search/Results?lookfor=&type=AllFields&botprotect="
                    continue
                elif retrieve_sign == "IxBib":
                    # "https://bible.ixtheo.de/Search/Results?lookfor=&type=AllFields&botprotect="
                    continue
                elif retrieve_sign == "KALDI/DaKaR":
                    # "https://churchlaw.ixtheo.de/Search/Results?lookfor=&type=AllFields&botprotect="
                    continue
                elif retrieve_sign in skip:
                    continue
                else:
                    print('Unbekanntes Abrufzeichen', retrieve_sign)
                    continue
                # print('http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.exk%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml')
                data = filedata.read()
                xml_soup = BeautifulSoup(data, "lxml")

                '''if xml_soup.find('datafield', tag='036F'):
                    if xml_soup.find('datafield', tag='036F').find('subfield', code='R'):
                        if xml_soup.find('datafield', tag='036F').find('subfield', code='R').text[1] == 'b':
                            print(ppn)
                            is_item_of_series = True
                if is_item_of_series:
                    items_of_series_data.append(row)
                else:
                    try:
                        filedata = urllib.request.urlopen('http://sru.k10plus.de/'
                                                          'opac-de-627?version=1.1&operation=searchRetrieve&query=pica.1049%3D' + ppn + '+and+pica.1045%3Drel-nt+and+pica.1001%3Db&maximumRecords=10&recordSchema=picaxml')
                        data = filedata.read()
                        xml_soup = BeautifulSoup(data, "lxml")
                        number_of_articles = xml_soup.find('zs:numberofrecords')
                        articles = False
                        if number_of_articles:
                            if int(number_of_articles.text) > 0:
                                int_number_of_articles = int(number_of_articles.text) - 1
                                row.insert(1, str(int_number_of_articles))
                                data_of_items_with_articles.append(row)
                                articles = True
                        if not articles:
                            selected_data.append(row)
                    except:
                        selected_data.append(row)'''
            except:
                request_not_successfull.append(row)
    #if not stop_processing:
    max_column = max(len(selected_data), len(items_of_series_data)) + 1
    extent_string = 'A1:V' + str(max_column)
    workbook = xlsxwriter.Workbook(
        path_to_directory + '/' + file_name.replace('.csv', '') + '_Stücktitel_bearbeitet.xlsx')
    first_worksheet = workbook.add_worksheet()
    first_worksheet.add_table(extent_string,
                              {'data': selected_data, 'style': 'Table Style Light 11', 'columns': header})
    second_worksheet = workbook.add_worksheet('Stücktitel')
    second_worksheet.add_table(extent_string,
                               {'data': items_of_series_data, 'style': 'Table Style Light 11', 'columns': header})
    third_worksheet = workbook.add_worksheet('PPNs mit Aufsätzen')
    third_worksheet.add_table(extent_string,
                               {'data': data_of_items_with_articles, 'style': 'Table Style Light 11', 'columns': header_for_items_with_articles})
    fourth_worksheet = workbook.add_worksheet('SRU-Abfrage gescheitert')
    fourth_worksheet.add_table(extent_string,
                              {'data': request_not_successfull, 'style': 'Table Style Light 11',
                               'columns': header})
    workbook.close()


'''def handle_click():
    file_name = entry.get()
    files = os.listdir(path_to_directory)
    if file_name in files:
        convert_excel(file_name)
    else:
        if file_name + '.csv' in files:
            file_name = file_name + '.csv'
            convert_excel(file_name)
        else:
            button.configure(bg="red")
            button.configure(text="Die angegebene CSV-Datei exisitert nicht.")
            window.update_idletasks()
            time.sleep(5)
        sys.exit()
    window.quit()'''


'''window = tk.Tk()
entry = tk.Entry(fg="black", bg="white", width=40)
label = tk.Label(text="Dateiname:", width=20)
button = tk.Button(text="Bestätigen",
                   width=30,
                   height=2,
                   bg="#7FFF00",
                   fg="black",
                   command=handle_click
                   )
quit_button = tk.Button(text="Beenden",
                        width=30,
                        height=2,
                        bg="#BEBEBE",
                        fg="black",
                        command=close_window
                        )
label.grid(column=0, row=0)   # grid dynamically divides the space in a grid
entry.grid(column=1, row=0)
button.grid(column=0, row=1, sticky="nsew")
quit_button.grid(column=1, row=1, sticky="nsew")

window.mainloop()'''

if __name__ == '__main__':
    convert_excel('Statistik_Abrufzeichen_8002_0575_2021_NEU.csv')
