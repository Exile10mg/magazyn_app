# Dokumentacja MagazynApp

## Opis programu

MagazynApp to prosty system magazynowy napisany w języku Python przy użyciu biblioteki Tkinter do tworzenia interfejsu graficznego. Program umożliwia użytkownikowi logowanie, wczytywanie danych z plików Excel dotyczących produktów (np. pomp, turbosprężarek, wtryskiwaczy) oraz porównywanie tych danych.

## Funkcje

### Logowanie

Po uruchomieniu programu, użytkownik może zalogować się, podając odpowiedni login i hasło. W obecnej implementacji poprawne dane to "admin" jako login i "admin" jako hasło.

### Wczytywanie danych

Program umożliwia wczytanie danych z dwóch plików Excel:
- **Dane Excel:** Pliki zawierające informacje o produktach, takie jak numer seryjny, ilość, itp.
- **Dane Subiekt:** Pliki zawierające informacje z systemu Subiekt, również o numerze seryjnym i ilości.

### Porównywanie danych

Po wczytaniu danych, użytkownik może porównać je, co skutkuje znalezieniem różnic między danymi z pliku Excel a danymi z systemu Subiekt. Różnice obejmują różnice w ilościach produktów.

### Wyświetlanie różnic

Program pozwala na wyświetlenie znalezionych różnic w formie nowego okna, które zawiera numery seryjne produktów, ilość z pliku Excel oraz ilość z systemu Subiekt.

### Eksport różnic

Użytkownik może eksportować znalezione różnice do nowego pliku Excel, który zawiera informacje o numerach seryjnych produktów oraz różnicach ilości między danymi z pliku Excel a danymi z systemu Subiekt.

## Struktura programu

Program składa się z dwóch klas:
- **MagazynApp:** Klasa główna, odpowiedzialna za interfejs użytkownika, logowanie, oraz otwieranie nowego okna dla porównania danych.
- **NewMagazynApp:** Klasa obsługująca nowe okno, w którym użytkownik może wybierać opcje porównania dla konkretnego rodzaju produktów.

## Wymagania

Aby uruchomić program, wymagane są następujące biblioteki:
- `tkinter`
- `ttk`
- `font`
- `messagebox`
- `filedialog`
- `pandas`
- `openpyxl`
- `Counter`
- `subprocess`
- `os`
- `PIL` (Pillow)

## Uruchomienie programu

Aby uruchomić program, należy uruchomić plik `magazyn.py`. Po uruchomieniu pojawi się okno logowania, gdzie użytkownik wprowadza dane.

## Autor

Program został stworzony przez Mike Boro dla Dakro Bosch Service. © 2023 Dakro Bosch Service.
