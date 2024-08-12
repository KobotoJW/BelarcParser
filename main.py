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

def get_pomieszczenie(path):
    return path.split("\\")[1]

def get_nazwa_komputera(soup):
    try:
        return soup.find('th', string='Computer Name:').find_next_sibling('td').text
    except AttributeError:
        try:
            return soup.find('th', string='Nazwa komputera:').find_next_sibling('td').text
        except AttributeError:
            return "Błąd"

def get_sprzet(soup):
    for table in soup.find_all('table'):
        caption = table.find('caption')
        if caption and caption.text.strip() in ['System Model', 'Model systemu']:
            return table.find('td').contents[0].strip() if caption.text.strip() == 'System Model' else table.find('td').contents[2].strip("systemu :")
    return ""

def get_procesor(soup):
    for table in soup.find_all('table'):
        caption = table.find('caption')
        if caption and caption.text.strip() in ['Processor a', 'Procesor a']:
            return table.find('td').contents[0].strip().replace("gigahertz", "GHz").replace("gigaherców", "GHz")
    return ""

def get_ram(soup):
    for table in soup.find_all('table'):
        caption = table.find('caption')
        if caption and caption.text.strip() in ['Memory Modules c,d', 'Moduły pamięci c,d']:
            ram_text = int(table.find('td').contents[0].strip('Megabytes Usable Installed Memory \n \t').strip('Megabajty Użyteczne zainstalowane gniazdo pamięci'))
            return f"{(ram_text / 1024).__round__(0)} GB"
    return ""

def get_dysk(soup):
    for table in soup.find_all('table'):
        caption = table.find('caption')
        if caption and caption.text.strip() in ['Drives', 'Dyski']:
            dysk_text = table.find('td').contents
            for elem in dysk_text:
                if "Hard drive" in elem or "Dysk twardy" in elem:
                    return elem.strip().split("--")[0]
    return ""

def get_windows(soup):
    for table in soup.find_all('table'):
        caption = table.find('caption')
        if caption and 'Software Licenses' in caption.text:
            win_text = table.find('tr').contents[1].text.split("\n")
            for elem in win_text:
                if 'Microsoft - Windows' in elem:
                    win_text = elem.strip('Microsoft - Windows').split(" ")
                    win_text = [item for item in win_text if not item.startswith("(x") and item != '(Key:']
                    win_text[-1] = win_text[-1].replace(')e', '')
                    return ' '.join(win_text[:-1]), win_text[-1]
    return "", ""

def get_office(soup):
    for table in soup.find_all('table'):
        caption = table.find('caption')
        if caption and 'Software Licenses' in caption.text:
            office_text = table.find('tr').contents[1].text.split("\n")
            for elem in office_text:
                if 'Microsoft - Office' in elem:
                    office_text = elem.strip('Microsoft - Office').split(" ")
                    office_text = [item for item in office_text if item != '(Key:']
                    office_text[-1] = office_text[-1].replace(')', '')
                    return office_text[0], office_text[2]
    return "Brak", ""

def get_antywirus(soup):
    for table in soup.find_all('table'):
        caption = table.find('caption')
        if caption and ('Virus Protection' in caption.text.strip() or 'Ochrona przed wirusami' in caption.text.strip()):
            return table.find('td').contents[1].text.strip().split("\n")[0]
    return ""

def get_mac(soup):
    for table in soup.find_all('table'):
        caption = table.find('caption')
        if caption and caption.text.strip() in ['Communications', 'Komunikacji']:
            mac_text = table.find('td').contents[3].text.split("\n")
            for elem in mac_text:
                if "Physical\xa0Address" in elem:
                    return ':'.join([item for item in elem.strip().split(":")[1:]])
    return ""

def parse_html(path):
    with open(path, 'r', encoding='utf-8') as file:
        content = file.read()
        soup = BeautifulSoup(content, 'html.parser')

        pomieszczenie = get_pomieszczenie(path)
        nazwa_komputera = get_nazwa_komputera(soup)
        sprzet = get_sprzet(soup)
        procesor = get_procesor(soup)
        ram = get_ram(soup)
        dysk = get_dysk(soup)
        windows, klucz_windows = get_windows(soup)
        office, klucz_office = get_office(soup)
        antywirus = get_antywirus(soup)
        mac = get_mac(soup)

        return [0, pomieszczenie, nazwa_komputera, sprzet, procesor, ram, dysk, windows, klucz_windows, office, klucz_office, antywirus, "", mac]
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
