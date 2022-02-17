import os
import xlsxwriter
import csv
import urllib.request
from bs4 import BeautifulSoup
import re
import pandas
from _datetime import datetime
import matplotlib.pyplot as plot
import matplotlib.ticker

path_to_directory = '//sn00.zdv.uni-tuebingen.de/ZE020110/FID-Projekte/Statistik/Statistik_automatisiert/'
path_to_directory_pandas = '//sn00.zdv.uni-tuebingen.de/ZE020110/FID-Projekte/Statistik/Statistik_automatisiert/'
# Angabe der Spaltenbezeichnungen
char_table = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']
# Angabe der Zellen, die fusioniert werden sollen; funktioniert derzeit nicht, deshalb nicht verwendet
merge_ranges = {'{0}2:{0}7': 'Datensatztypen', '{0}9:{0}15': 'Projekte', '{0}17:{0}20': 'Produktionsverfahren',
                '{0}21:{0}23': 'Abzugskennzeichen', '{0}25:{0}39': 'Kooperationen', '{0}41:{0}46': 'SSGs', '{0}47:{0}50': 'WebDatenbanken'}


def convert_excel(file_name):
    timestamp = datetime.now()
    timestamp = timestamp.strftime("%d.%m.%Y")
    if 'Statistik_Abrufzeichen_' + timestamp + '.xlsx' not in os.listdir(path_to_directory):
        merge_cells = {}
        # einlesen der Datei mit dem letzten Stand
        read_file = pandas.read_excel(path_to_directory_pandas + 'Statistik_Abrufzeichen_aktuell.xlsx')
        # Umwandeln der Datei in eine CSV-Datei
        read_file.to_csv(path_to_directory_pandas + 'Statistik_Abrufzeichen_aktuell.csv', index=None, header=True)
        # Einlesen und Parsen der generierten CSV-Datei
        csv_file = open(path_to_directory + '/' + file_name, 'r', encoding='utf-8')
        csv_reader = csv.reader(csv_file, delimiter=',', quotechar='"')
        rows = [row for row in csv_reader]
        # Erstellen einer Liste, um die Zeilen der neuen Datei abzulegen
        statistics_data = []
        request_not_successfull = []
        header = []
        # Einsetzen eines Zählers für die Zeilen der CSV-Datei
        row_nr = 0
        last_column_nr = False
        for row in rows:
            row_nr += 1
            if row_nr == 1:
                first_date = False
                column_nr = 0
                for column in row:
                    # Festlegung der letzten Spalte, die statistische Daten enthält
                    if re.findall(r'\d{2}\.\d{2}\.\d{4}', column):
                        first_date = True
                    if first_date == True and not re.findall(r'\d{2}\.\d{2}\.\d{4}', column):
                        last_column_nr = column_nr
                        break
                    column_nr += 1
                if not last_column_nr:
                    last_column_nr = column_nr
                # Spaltenüberschriften speichern
                header = [{'header': column} for column in row]
                # hinzufügen einer neuen Spalte mit dem aktuellen Datum
                rows[0].insert(last_column_nr, timestamp)
                header.insert(last_column_nr, {'header': timestamp})
                continue
            # Durchgang durch die Abrufzeichen
            retrieve_sign = row[1]
            # Festlegung der Abrufzeichen, die im Exemplarsatz stehen
            retrieve_signs_exk = ["ixau", "ixmo", "ixsb", "ixze", "ixzg", "ixzs", "ixzw", "ixzx", "ixrk", "rwzw", "rwzx",
                                  "rwrk", "ixzo", "zota", "imwa", "ixbt", "rwex", "bril", "gruy",
                                  "knix", "kn28", "mszo", "mszk", "msmi", "bsbo"]
            # Festlegung der Abrufzeichen, die im Titeldatensatz stehen
            retrieve_signs_cod = ["DTH5", "mteo", "mtex", "BIIN", "KALD", "GIRA", "DAKR", "AUGU", "MIKA", "redo"]
            # Festlegung restlicher Abrufzeichen
            retrieve_signs_sfk = ["1", "0", "6,22", "1 or 0", "1 or 0 or 6,22", "1 or 0 or 6,22 or mteo not mtex"]
            # Überspringen der Trennzeilen
            if retrieve_sign:
                try:
                    # Abruf der Ergebnismenge über die SRU-Schnittstelle
                    if retrieve_sign in retrieve_signs_exk:
                        filedata = urllib.request.urlopen('https://sru.bsz-bw.de/cbsx?version=1.1&operation=searchRetrieve&query=pica.exk%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml&x-username=s2304&x-password=3i1Q')
                        data = filedata.read()
                        xml_soup = BeautifulSoup(data, "lxml")
                        number_of_records = xml_soup.find('zs:numberofrecords').text
                    elif retrieve_sign in retrieve_signs_cod:
                        filedata = urllib.request.urlopen('https://sru.bsz-bw.de/cbsx?version=1.1&operation=searchRetrieve&query=pica.cod%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml&x-username=s2304&x-password=3i1Q')
                        data = filedata.read()
                        xml_soup = BeautifulSoup(data, "lxml")
                        number_of_records = xml_soup.find('zs:numberofrecords').text
                    elif retrieve_sign in retrieve_signs_sfk:
                        if 'or' not in retrieve_sign:
                            filedata = urllib.request.urlopen(
                            'https://sru.bsz-bw.de/cbsx?version=1.1&operation=searchRetrieve&query=pica.sfk%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml&x-username=s2304&x-password=3i1Q')
                            data = filedata.read()
                            xml_soup = BeautifulSoup(data, "lxml")
                            number_of_records = xml_soup.find('zs:numberofrecords').text
                        elif retrieve_sign == "1 or 0 or 6,22 or mteo not mtex":
                            filedata = urllib.request.urlopen("https://sru.bsz-bw.de/cbsx?version=1.1&operation=searchRetrieve&query=pica.sfk%3D0+or+pica.sfk%3D1+or+pica.sfk%3D6,22+or+pica.cod%3Dmteo+not+pica.cod%3Dmtex&maximumRecords=10&recordSchema=picaxml&x-username=s2304&x-password=3i1Q")
                            data = filedata.read()
                            xml_soup = BeautifulSoup(data, "lxml")
                            number_of_records = xml_soup.find('zs:numberofrecords').text
                        else:
                            retrieve_sign_list = retrieve_sign.split(' or ')
                            retrieve_sign = "+or+pica.sfk%3D".join(retrieve_sign_list)
                            filedata = urllib.request.urlopen(
                                'https://sru.bsz-bw.de/cbsx?version=1.1&operation=searchRetrieve&query=pica.sfk%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml&x-username=s2304&x-password=3i1Q')
                            data = filedata.read()
                            xml_soup = BeautifulSoup(data, "lxml")
                            number_of_records = xml_soup.find('zs:numberofrecords').text
                    # Abruf der Ergebnismenge aus den Web-Datenbanken
                    elif retrieve_sign == "IxTheo":
                        filedata = urllib.request.urlopen("https://www.ixtheo.de/Search/Results?lookfor=&type=AllFields&botprotect=")
                        data = filedata.read()
                        xml_soup = BeautifulSoup(data, "lxml")
                        number_of_records = re.findall(r'results of ([\d,]*) for search', xml_soup.find('div', class_='search-stats').text)
                        if number_of_records:
                            number_of_records = number_of_records[0].replace(',', '')
                    elif retrieve_sign == "RelBib":
                        filedata = urllib.request.urlopen("https://www.relbib.de/Search/Results?lookfor=&type=AllFields&botprotect=")
                        data = filedata.read()
                        xml_soup = BeautifulSoup(data, "lxml")
                        number_of_records = re.findall(r'results of ([\d,]*) for search',
                                                       xml_soup.find('div', class_='search-stats').text)
                        if number_of_records:
                            number_of_records = number_of_records[0].replace(',', '')
                    elif retrieve_sign == "IxBib":
                        filedata = urllib.request.urlopen("https://bible.ixtheo.de/Search/Results?lookfor=&type=AllFields&botprotect=")
                        data = filedata.read()
                        xml_soup = BeautifulSoup(data, "lxml")
                        number_of_records = re.findall(r'results of ([\d,]*) for search',
                                                       xml_soup.find('div', class_='search-stats').text)
                        if number_of_records:
                            number_of_records = number_of_records[0].replace(',', '')
                    elif retrieve_sign == "KALDI/DaKaR":
                        filedata = urllib.request.urlopen("https://churchlaw.ixtheo.de/Search/Results?lookfor=&type=AllFields&botprotect=")
                        data = filedata.read()
                        xml_soup = BeautifulSoup(data, "lxml")
                        number_of_records = re.findall(r'results of ([\d,]*) for search',
                                                       xml_soup.find('div', class_='search-stats').text)
                        if number_of_records:
                            number_of_records = number_of_records[0].replace(',', '')
                    elif retrieve_sign == "Seit August 2020 >2.300:":
                        number_of_records = ""
                    else:
                        print('Unbekanntes Abrufzeichen', retrieve_sign)
                        number_of_records = ""
                    # hinzufügen der neuen Daten zur jeweiligen Zeile
                    row.insert(last_column_nr, number_of_records)
                    statistics_data.append(row)
                except:
                    print('request not successfull:', row)
                    request_not_successfull.append(row)
            # Übernahme der Trennzeilen als leere Zeilen
            else:
                row.insert(last_column_nr, "")
                statistics_data.append(row)
        # Generierung der Daten für die grafische Aufbereitung
        # for row in rows:
            # print(row)
        if len(rows) > 0:
            # Erfassung der Zeitstempel für die x-Achse
            dates = [date[0:6] for date in rows[0][2:last_column_nr + 1]]
            # Erfassung und Zusammenfassung der Daten nach Kooperationen, Datentypen, Projekten, Datenbanken und Produktionsverfahren
            coops = {}
            datatypes = {}
            projects = {}
            databases = {}
            production_processes = {}
            knix_list = [float(num) for num in [row[2:last_column_nr + 1] for row in statistics_data if row[1] == 'knix'][0]]
            kn28_list = [float(num) for num in [row[2:last_column_nr + 1] for row in statistics_data if row[1] == 'kn28'][0]]
            kn_total = [str(sum(x)) for x in zip(knix_list, kn28_list)]
            rwzw_list = [float(num) for num in [row[2:last_column_nr + 1] for row in statistics_data if row[1] == 'rwzw'][0]]
            rwzx_list = [float(num) for num in [row[2:last_column_nr + 1] for row in statistics_data if row[1] == 'rwzx'][0]]
            rw_total = [str(sum(x)) for x in zip(rwzw_list, rwzx_list)]
            ixzw_list = [float(num) for num in [row[2:last_column_nr + 1] for row in statistics_data if row[1] == 'ixzw'][0]]
            ixzx_list = [float(num) for num in [row[2:last_column_nr + 1] for row in statistics_data if row[1] == 'ixzx'][0]]
            ix_total = [str(sum(x)) for x in zip(ixzw_list, ixzx_list)]
            projects['ixzw + ixzx'] = ix_total
            projects['rwzw + rwzx'] = rw_total
            coops['knix + kn28'] = kn_total
            for row in statistics_data:
                if row[1] in ['BIIN', 'KALD', 'GIRA', 'DAKR', 'AUGU', 'MIKA', 'redo', 'bsbo']:
                    coops[row[1]] = row[2:last_column_nr + 1]
                elif row[1] in ['ixau', 'ixmo', 'ixze', 'ixzg', 'ixzs']:
                    datatypes[row[1]] = row[2:last_column_nr + 1]
                elif row[1] in ['DTH5', 'ixrk', 'rwrk']:
                    projects[row[1]] = row[2:last_column_nr + 1]
                elif row[1] in ['IxTheo', 'RelBib', 'IxBib', 'KALDI/DaKaR']:
                    databases[row[1]] = row[2:last_column_nr + 1]
                elif row[1] in ['ixzo', 'zota', 'imwa', 'ixbt']:
                    production_processes[row[1]] = row[2:last_column_nr + 1]
            # Erstellen der Grafiken
            plot.rc('xtick', labelsize=10)
            plot.rc('ytick', labelsize=10)
            fig, axs = plot.subplots(3, 2, figsize=(20, 16)) # Gesamtgröße des Plots in Inch
            row_nr = -1
            data_nr = 0
            # Anlegen eines Dictionarys, um den Zuwachs seit der letzten Datenerfassung abzuspeichern
            growth = {'Kooperationen': {}, 'Produktionsverfahren': {}}
            for data_for_graphic in [datatypes, coops, projects, databases, production_processes]:
                if data_nr%2 == 0:
                    row_nr += 1
                for row in data_for_graphic:
                    try:
                        new_data = [int(float(re.sub(r'\.0$', ' ', number).replace('*', '').replace('.', ''))) for number in data_for_graphic[row]]
                        axs[row_nr, data_nr%2].get_yaxis().set_major_formatter(
                            matplotlib.ticker.FuncFormatter(lambda x, p: format(float(x), ',')))
                        if data_for_graphic == datatypes:
                            label = 'Datentypen'
                        elif data_for_graphic == coops:
                            label = 'Kooperationen'
                            growth['Kooperationen'][row] = new_data[-1] - new_data[-2]
                        elif data_for_graphic == projects:
                            label = 'Projekte'
                        elif data_for_graphic == production_processes:
                            growth['Produktionsverfahren'][row] = new_data[-1] - new_data[-2]
                            continue
                        else:
                            label = 'Datenbanken'
                        axs[row_nr, data_nr%2].set_title(label)
                        axs[row_nr, data_nr % 2].plot(dates, new_data, label=row)
                        axs[row_nr, data_nr % 2].legend()
                    except Exception as e:
                        print(e)
                        print(row)
                data_nr += 1
            # plotten der Grafiken in einem PDF-Dokument
            wedges, texts = axs[2, 0].pie([growth['Produktionsverfahren'][key] for key in growth['Produktionsverfahren']])
            axs[2, 0].legend(wedges, [key + ' (' + str(growth['Produktionsverfahren'][key]) + ')' for key in growth['Produktionsverfahren']], bbox_to_anchor=(-0.1, 1.),
                       fontsize=8)
            axs[2,0].set_title('Zuwächse nach (halb-)automatischen Produktionsverfahren')
            wedges, texts = axs[2, 1].pie([growth['Kooperationen'][key] if growth['Kooperationen'][key] >= 0 else 0 for key in growth['Kooperationen']])
            axs[2, 1].legend(wedges, [key + ' (' + str(growth['Kooperationen'][key]) + ')' for key in growth['Kooperationen']], bbox_to_anchor=(-0.1, 1.),
                             fontsize=8)
            axs[2, 1].set_title('Zuwächse nach Kooperationspartnern')
            pdf_filename = file_name.replace('aktuell.csv', timestamp).replace('.', '_')
            plot.savefig(path_to_directory + '/' + pdf_filename + '.pdf')
        # Excel-Tabelle mit akutellem Zeitstempel generieren
        max_column = len(statistics_data) + 1
        extent_string = 'A1:V' + str(max_column)
        for new_filename in [file_name.replace('aktuell.csv', timestamp), file_name.replace('.csv', '')]:
            workbook = xlsxwriter.Workbook(
                path_to_directory + '/' + new_filename + '.xlsx', {'strings_to_numbers': True})
            first_worksheet = workbook.add_worksheet()
            '''
            merge_format = workbook.add_format({
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'fg_color': 'yellow'})
            for merge_range in merge_ranges:
                first_worksheet.merge_range(merge_range.format(last_column_char), merge_ranges[merge_range], merge_format)
            '''
            # es scheint nicht möglich zu sein, eine Tabelle zu generieren und gleichzeitig Zellen zu mergen.

            first_worksheet.add_table(extent_string,
                                      {'data': statistics_data, 'style': 'Table Style Medium 1', 'columns': header})
            workbook.close()


if __name__ == '__main__':
    convert_excel('Statistik_Abrufzeichen_aktuell.csv')
