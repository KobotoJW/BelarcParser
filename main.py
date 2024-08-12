import pyexcel_ods3
import os
from bs4 import BeautifulSoup

def walklevel(path, depth = 1):
    if depth < 0:
        for root, dirs, files in os.walk(path):
            yield root, dirs[:], files
        return
    elif depth == 0:
        return
    base_depth = path.rstrip(os.path.sep).count(os.path.sep)
    for root, dirs, files in os.walk(path):
        yield root, dirs[:], files
        cur_depth = root.count(os.path.sep)
        if base_depth + depth <= cur_depth:
            del dirs[:]

def create_sheet():
    return {"Komputery": [["LP", "Pomieszczenie", "Nazwa komputera", "Sprzęt", "Procesor", "RAM", "Dysk", "Windows", "Klucz windows", "Office", "Klucz office", "Antywirus", "Do kiedy", "MAC"]]}

def generate_filename():
    if not os.path.exists("Spis_komputerow.ods"):
        return "Spis_komputerow.ods"
    else:
        i = 1
        while os.path.exists("Spis_komputerow" + str(i) + ".ods"):
            i += 1
        return "Spis_komputerow" + str(i) + ".ods"

def write_to_ods(buff, file_name):
    pyexcel_ods3.save_data(file_name, buff)

def extract_html_paths():
    htmls = []
    for root, dirs, files in walklevel(".", 1):
        for file in files:
            if file.endswith(".html") or file.endswith(".htm"):
                htmls.append(os.path.join(root, file))

    return htmls

def parse_html(path):
    pomieszczenie = ""
    nazwa_komputera = ""
    sprzet = ""
    procesor = ""
    ram = ""
    dysk = ""
    windows = ""
    klucz_windows = ""
    office = ""
    klucz_office = ""
    antywirus = ""
    do_kiedy = ""
    mac = ""

    with open(path, 'r', encoding='utf-8') as file:
        content = file.read()
        soup = BeautifulSoup(content, 'html.parser')
        tables = soup.find_all('table')
        # for table in tables:
        #     print(table.find('caption'))

        pomieszczenie = path.split("\\")[1]
        ##########################################
        if (soup.find('th', string = 'Computer Name:') != None):
            nazwa_komputera = soup.find('th', string = 'Computer Name:').find_next_sibling('td').text
        else:
            try:
                nazwa_komputera = soup.find('th', string = 'Nazwa komputera:').find_next_sibling('td').text
            except AttributeError:
                nazwa_komputera = "Błąd"
        ##########################################
        system_model_text = ""
        for table in tables:
            caption = table.find('caption')
            if caption and caption.text.strip() == 'System Model':
                system_model_text = table.find('td').contents[0].strip()
                break
            elif caption and caption.text.strip() == 'Model systemu':
                system_model_text = table.find('td').contents[2].strip("systemu :")
                break
        sprzet = system_model_text
        ##########################################
        processor_text = ""
        for table in tables:
            caption = table.find('caption')
            if caption and caption.text.strip() == 'Processor a':
                processor_text = table.find('td').contents[0].strip().replace("gigahertz", "GHz")
                break
            elif caption and caption.text.strip() == 'Procesor a':
                processor_text = table.find('td').contents[0].strip().replace("gigaherców", "GHz")
                break
        if 'processor_text' in locals():
            procesor = processor_text
        ##########################################
        ram_text = ""
        for table in tables:
            caption = table.find('caption')
            if caption and caption.text.strip() == 'Memory Modules c,d':
                ram_text = int(table.find('td').contents[0].strip('Megabytes Usable Installed Memory \n \t'))
                break
            elif caption and caption.text.strip() == 'Moduły pamięci c,d':
                ram_text = int(table.find('td').contents[0].strip('Megabajty Użyteczne zainstalowane gniazdo pamięci'))
                break
        if ram_text != "":
            ram_text = (ram_text / 1024).__round__(0)
            ram_text = str(ram_text) + " GB"
        ram = ram_text
        ##########################################
        dysk_text = ""
        for table in tables:
            caption = table.find('caption')
            if caption and caption.text.strip() == 'Drives':
                dysk_text = table.find('td').contents
                for elem in dysk_text:
                    if "Hard drive" in elem:
                        dysk_text = elem.strip().split("--")[0]
                break
            elif caption and caption.text.strip() == 'Dyski':
                dysk_text = table.find('td').contents
                for elem in dysk_text:
                    if "Dysk twardy" in elem:
                        dysk_text = elem.strip().split("--")[0]
                break
        if dysk_text:
            dysk = dysk_text
        ##########################################
        win_text = ""
        office_text = ""
        for table in tables:
            caption = table.find('caption')
            if caption and 'Software Licenses' in caption.text:
                win_text = table.find('tr').contents[1].text
                break
        win_text = win_text.split("\n")
        office_text = win_text
        for elem in win_text:
            if 'Microsoft - Windows' in elem:
                win_text = elem.strip('Microsoft - Windows').split(" ")
        for i in range(len(win_text)):
            if win_text[i].startswith("(x"):
                win_text[i] = win_text[i][1:4]
        try:
            win_text.remove('(Key:')
        except ValueError:
            pass
        win_text[-1] = win_text[-1].replace(')e', '')
        windows = ' '.join([item for item in win_text[0:-1]])
        if win_text[-1]:
            klucz_windows = win_text[-1]
        ##########################################
        for elem in office_text:
            if 'Microsoft - Office' in elem:
                office_text = elem.strip('Microsoft - Office').split(" ")
                break
            if 'Microsoft - Office' not in office_text:
                office_text = 'Brak'
        if office_text != 'Brak':
            office_text.remove('(Key:')
            office_text[-1].replace(')', '')
            office_text.pop(1)
            klucz_office = office_text[1]
            office_text = office_text[0]
        office = office_text
        ##########################################
        antivirus_text = ""
        for table in tables:
            caption = table.find('caption')
            if caption and 'Virus Protection' in caption.text.strip():
                antivirus_text = table.find('td').contents[1].text.strip().split("\n")[0]
                break
            elif caption and 'Ochrona przed wirusami' in caption.text.strip():
                antivirus_text = table.find('td').contents[1].text.strip().split("\n")[0]
                break
        if antivirus_text:
            antywirus = antivirus_text
        ##########################################
        do_kiedy = ""
        ##########################################
        mac_text = ""
        for table in tables:
            caption = table.find('caption')
            if caption and caption.text.strip() == 'Communications':
                mac_text = table.find('td').contents[3].text
                mac_text = mac_text.split("\n")
                for elem in mac_text:
                    if "Physical\xa0Address" in elem:
                        mac_text = ':'.join([item for item in elem.strip().split(":")[1:]])
                        break
                break
            elif caption and caption.text.strip() == 'Komunikacji':
                mac_text = table.find('td').contents[3].text
                mac_text = mac_text.split("\n")
                for elem in mac_text:
                    if "Physical\xa0Address" in elem:
                        mac_text = ':'.join([item for item in elem.strip().split(":")[1:]])
                        break
        if mac_text:
            mac = mac_text

        #print(mac)
        return [0, pomieszczenie, nazwa_komputera, sprzet, procesor, ram, dysk, windows, klucz_windows, office, klucz_office, antywirus, do_kiedy, mac]

def main():
    lp = 1
    buff = create_sheet()
    for html in extract_html_paths():
       row = parse_html(html)
       row[0] = lp
       lp += 1
       buff["Komputery"].append(row)
    write_to_ods(buff, generate_filename())
    return 0

if __name__ == "__main__":
    main()
