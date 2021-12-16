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


def convert_excel(file_name):
    timestamp = datetime.now()
    timestamp = timestamp.strftime("%d.%m.%Y")
    if 'Statistik_Abrufzeichen_' + timestamp + '.xlsx' not in os.listdir(path_to_directory):
        read_file = pandas.read_excel(path_to_directory_pandas + 'Statistik_Abrufzeichen_aktuell.xlsx')
        read_file.to_csv(path_to_directory_pandas + 'Statistik_Abrufzeichen_aktuell.csv', index=None, header=True)
        csv_file = open(path_to_directory + '/' + file_name, 'r', encoding='utf-8')
        csv_reader = csv.reader(csv_file, delimiter=',', quotechar='"')
        rows = [row for row in csv_reader]
        items_of_series_data = []
        selected_data = []
        request_not_successfull = []
        row_nr = 0
        header = []
        total_rows = len(rows)
        units = int(total_rows/100)
        counter = 0
        total = 0
        last_column_nr = False
        for row in rows:
            '''if stop_processing:
                break'''
            row_nr += 1
            if row_nr == 1:
                first_date = False
                column_nr = 0
                for column in row:
                    if re.findall(r'\d{2}\.\d{2}\.\d{4}', column):
                        first_date = True
                    if first_date == True and not re.findall(r'\d{2}\.\d{2}\.\d{4}', column):
                        last_column_nr = column_nr
                        break
                    column_nr += 1
                header = [{'header': column} for column in row]
                rows[0].insert(last_column_nr, timestamp)
                header.insert(last_column_nr, {'header': timestamp})
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
            retrieve_signs_exk = ["ixau", "ixmo", "ixsb", "ixze", "ixzg", "ixzs", "ixzw", "ixzx", "ixrk", "rwzw", "rwzx",
                                  "rwrk", "ixzo", "zota", "imwa", "ixbt", "rwex", "bril", "gruy",
                                  "knix", "kn28", "mszo", "mszk", "msmi", "redo", "bsbo"]
            retrieve_signs_cod = ["DTH5", "mteo", "mtex", "BIIN", "KALD", "GIRA", "DAKR", "AUGU", "MIKA"]
            retrieve_signs_sfk = ["1", "0", "6,22", "1 or 0", "1 or 0 or 6,22", "1 or 0 or 6,22 or mteo not mtex"]
            if retrieve_sign:
                total += 1
                try:
                    if retrieve_sign in retrieve_signs_exk:
                        filedata = urllib.request.urlopen('http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.exk%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml')
                        data = filedata.read()
                        xml_soup = BeautifulSoup(data, "lxml")
                        number_of_records = xml_soup.find('zs:numberofrecords').text
                    elif retrieve_sign in retrieve_signs_cod:
                        filedata = urllib.request.urlopen('http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.cod%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml')
                        data = filedata.read()
                        xml_soup = BeautifulSoup(data, "lxml")
                        number_of_records = xml_soup.find('zs:numberofrecords').text
                    elif retrieve_sign in retrieve_signs_sfk:
                        if 'or' not in retrieve_sign:
                            filedata = urllib.request.urlopen(
                            'http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.sfk%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml')
                            data = filedata.read()
                            xml_soup = BeautifulSoup(data, "lxml")
                            number_of_records = xml_soup.find('zs:numberofrecords').text
                        elif retrieve_sign == "1 or 0 or 6,22 or mteo not mtex":
                            filedata = urllib.request.urlopen("http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.sfk%3D0+or+pica.sfk%3D1+or+pica.sfk%3D6,22+or+pica.cod%3Dmteo+not+pica.cod%3Dmtex&maximumRecords=10&recordSchema=picaxml")
                            data = filedata.read()
                            xml_soup = BeautifulSoup(data, "lxml")
                            number_of_records = xml_soup.find('zs:numberofrecords').text
                        else:
                            retrieve_sign_list = retrieve_sign.split(' or ')
                            retrieve_sign = "+or+pica.sfk%3D".join(retrieve_sign_list)
                            filedata = urllib.request.urlopen(
                                'http://sru.k10plus.de/opac-de-627?version=1.1&operation=searchRetrieve&query=pica.sfk%3D' + retrieve_sign + '&maximumRecords=10&recordSchema=picaxml')
                            data = filedata.read()
                            xml_soup = BeautifulSoup(data, "lxml")
                            number_of_records = xml_soup.find('zs:numberofrecords').text
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
                        continue
                    # number_of_records = re.sub(r'(?<!^)(?=(\d{3})+$)', r'.', number_of_records)
                    row.insert(last_column_nr, number_of_records)
                    selected_data.append(row)
                except:
                    request_not_successfull.append(row)
        # Grafiken zusammenstellen:
        # Datensatztypen weglassen
        # brill + gruy nicht übernehmen
        # alle Projekte in eine Grafik:
            # iszw + ixzx in eine Linie
            # rwzw + rwzx in eine Linie
            # ixrk + rwrk in eine Linie
            # DTH5 in eine Linie
        # alle Produktionsverfahren in eine Grafik
        # alle Kooperationen in eine Grafik:
            # knix und kn28 in eine Linie
            # alle ms** weglassen
        # alle Webdatenbanken in eine Grafik
        # SSG-Abfragen weglassen
        last_column_nr += 1
        if len(rows) > 0:
            data_1 = [date[0:6] for date in rows[0][3:last_column_nr]]
            coops = {}
            datatypes = {}
            projects = {}
            databases = {}
            production_processes = {}
            knix_list = [int(num) for num in [row[3:last_column_nr] for row in selected_data if row[1] == 'knix'][0]]
            kn28_list = [int(num) for num in [row[3:last_column_nr] for row in selected_data if row[1] == 'kn28'][0]]
            kn_total = [str(sum(x)) for x in zip(knix_list, kn28_list)]
            rwzw_list = [int(num) for num in [row[3:last_column_nr] for row in selected_data if row[1] == 'rwzw'][0]]
            rwzx_list = [int(num) for num in [row[3:last_column_nr] for row in selected_data if row[1] == 'rwzx'][0]]
            rw_total = [str(sum(x)) for x in zip(rwzw_list, rwzx_list)]
            ixzw_list = [int(num) for num in [row[3:last_column_nr] for row in selected_data if row[1] == 'ixzw'][0]]
            ixzx_list = [int(num) for num in [row[3:last_column_nr] for row in selected_data if row[1] == 'ixzx'][0]]
            ix_total = [str(sum(x)) for x in zip(ixzw_list, ixzx_list)]
            projects['ixzw + ixzx'] = ix_total
            projects['rwzw + rwzx'] = rw_total
            coops['knix + kn28'] = kn_total
            for row in selected_data:
                if row[1] in ['BIIN', 'KALD', 'GIRA', 'DAKR', 'AUGU', 'MIKA', 'redo', 'bsbo']:
                    coops[row[1]] = row[3:last_column_nr]
                elif row[1] in ['ixau', 'ixmo', 'ixze', 'ixzg', 'ixzs']:
                    datatypes[row[1]] = row[3:last_column_nr]
                elif row[1] in ['DTH5', 'ixrk', 'rwrk']:
                    projects[row[1]] = row[3:last_column_nr]
                elif row[1] in ['IxTheo', 'RelBib', 'IxBib', 'KALDI/DaKaR']:
                    databases[row[1]] = row[3:last_column_nr]
                elif row[1] in ['ixzo', 'zota', 'imwa', 'ixbt']:
                    production_processes[row[1]] = row[3:last_column_nr]
            plot.rc('xtick', labelsize=10)
            plot.rc('ytick', labelsize=10)
            fig, axs = plot.subplots(3, 2, figsize=(20, 16)) # Gesamtgröße des Plots in Inch
            row_nr = -1
            data_nr = 0
            growth = {'Kooperationen': {}, 'Produktionsverfahren': {}}
            for data_for_graphic in [datatypes, coops, projects, databases, production_processes]:
                # plot.tight_layout()
                if data_nr%2 == 0:
                    row_nr += 1
                for row in data_for_graphic:
                    try:
                        new_data = [int(number.replace('*', '').replace('.', '')) for number in data_for_graphic[row]]
                        axs[row_nr, data_nr%2].get_yaxis().set_major_formatter(
                            matplotlib.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
                        # axs[row_nr, data_nr % 2].get_xaxis().xticks(fontsize=8)
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
                        axs[row_nr, data_nr % 2].plot(data_1, new_data, label=row)
                        axs[row_nr, data_nr % 2].legend()

                    except Exception as e:
                        print(e)
                data_nr += 1
            print(growth)

            wedges, texts = axs[2, 0].pie([growth['Produktionsverfahren'][key] for key in growth['Produktionsverfahren']])
            axs[2, 0].legend(wedges, [key + ' (' + str(growth['Produktionsverfahren'][key]) + ')' for key in growth['Produktionsverfahren']], bbox_to_anchor=(-0.1, 1.),
                       fontsize=8)
            axs[2,0].set_title('Zuwächse nach Produktionsverfahren')
            wedges, texts = axs[2, 1].pie([growth['Kooperationen'][key] if growth['Kooperationen'][key] >= 0 else 0 for key in growth['Kooperationen']])
            axs[2, 1].legend(wedges, [key + ' (' + str(growth['Kooperationen'][key]) + ')' for key in growth['Kooperationen']], bbox_to_anchor=(-0.1, 1.),
                             fontsize=8)
            axs[2, 1].set_title('Zuwächse nach Kooperationspartnern')

            # plot.show()
            pdf_filename = file_name.replace('aktuell.csv', timestamp).replace('.', '_')
            plot.savefig(path_to_directory + '/' + pdf_filename + '.pdf')
        max_column = max(len(selected_data), len(items_of_series_data)) + 1
        extent_string = 'A1:V' + str(max_column)
        for new_filename in [file_name.replace('aktuell.csv', timestamp), file_name.replace('.csv', '')]:
            workbook = xlsxwriter.Workbook(
                path_to_directory + '/' + new_filename + '.xlsx', {'strings_to_numbers': True})
            first_worksheet = workbook.add_worksheet()

            first_worksheet.add_table(extent_string,
                                      {'data': selected_data, 'style': 'Table Style Light 11', 'columns': header})
            workbook.close()


if __name__ == '__main__':
    convert_excel('Statistik_Abrufzeichen_aktuell.csv')
