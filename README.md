# BelarcParser

Parser do plików html/htm tworzonych przez program Belarc Advisor (Belarc Advisor Computer Profile).

Program tworzy arkusz OpenCalc (.ods) zawierający kolumny: "Pomieszczenie", "Nazwa komputera", "Sprzęt", "Procesor", "RAM", "Dysk", "Windows", "Klucz windows", "Office", "Klucz office", "Antywirus", "MAC".

## Użycie
Każdy plik powinien znajdować się w swoim własnym folderze nazwanym nazwą pomieszczenia. Przykład:
```
Folder
|
|- Sala1
|  |-- Belarc Advisor Computer Profile.html 
|
|- Sala2
|  |-- Belarc Advisor Computer Profile.html 
|
|- main.py OR belarc_parser-v0.5.exe
```

Taki układ spowoduje utworzenie pliku z 2 wierszami z danymi komputerów z nazwami Sala1 i Sala2 w kolumnie Pomieszczenie.

Skrypt uruchamiamy klasycznie: `python main.py`.
Plik wynikowy pojawi się w folderze na tym samym poziomie, z którego uruchomiono skrypt.

Plik .exe można stworzyć poleceniem `pyinstaller -F --hidden-import pyexcel_io.writers main.py`. (Teraz dostępny do pobrania w Releases)

Jeśli używamy pliku .exe umieszczamy go w tym samym miejscu zamiast kodu.
