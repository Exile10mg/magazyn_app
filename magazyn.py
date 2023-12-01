import tkinter as tk
from tkinter import ttk, font
from tkinter import messagebox
from tkinter import filedialog
from tkinter import font
import pandas as pd
from tkinter import ttk, simpledialog
import openpyxl
from collections import Counter
import subprocess
import os
from PIL import Image, ImageTk

class MagazynApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Magazyn")
        # Dodaj ikonę aplikacji
        self.master.iconbitmap("logo_ico.ico")  # Zastąp 'logo.ico' nazwą twojego pliku ikony

        # Wczytaj obraz z pliku PNG
        image = Image.open("logo.png")  # Zastąp 'logo.png' nazwą twojego pliku PNG
        photo = ImageTk.PhotoImage(image)

        # Utwórz etykietę z obrazem
        logo = ttk.Label(self.master, image=photo)
        logo.photo = photo  # Zapobiegnij zniknięciu obrazu z pamięci
        logo.pack(pady=1)  # Zwiększyłem pady, aby logo było wyżej

        # Czcionka
        custom_font = font.nametofont("TkDefaultFont")
        custom_font.configure(size=custom_font.cget("size"))

        # Styl przycisku
        style = ttk.Style()
        style.configure("TButton", padding=10, relief="flat", background="#4CAF50", foreground="black", font=custom_font)

        # Panel logowania
        self.login_frame = ttk.Frame(self.master)
        self.login_frame.pack(pady=1)  # Zwiększyłem pady, aby panel logowania było niżej

        login_label = ttk.Label(self.login_frame, text="Login:")
        login_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entry_login = tk.Entry(self.login_frame, width=15)
        self.entry_login.grid(row=0, column=1, pady=5)

        password_label = ttk.Label(self.login_frame, text="Hasło:")
        password_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_password = tk.Entry(self.login_frame, show="*", width=15)
        self.entry_password.grid(row=1, column=1, pady=5)

        # Przycisk zaloguj
        login_button = ttk.Button(self.login_frame, text="Zaloguj", command=self.login_info)
        login_button.grid(row=2, column=0, columnspan=2, pady=10)

        # Ustawienia dostosowujące rozmiar okna do treści
        self.master.update()
        self.master.geometry("{}x{}".format(self.master.winfo_reqwidth(), self.master.winfo_reqheight()))

        # Stopka
        label_opcje = ttk.Label(self.master, text="© 2023 Dakro Bosch Service Autor: Mike Boro", font=("Arial", 8))
        label_opcje.pack(pady=2, padx=2)

    def login_info(self):
        login = self.entry_login.get()
        password = self.entry_password.get()
        if login == "admin" and password == "admin":
            messagebox.showinfo("Logowanie", "Zalogowano pomyślnie!")

            # Tutaj dodano nowe okno i zamknięcie poprzedniego
            self.open_new_window()
            self.master.destroy()
        else:
            messagebox.showinfo("Logowanie", "Niepoprawne dane!")

    def open_new_window(self):
        new_window = tk.Tk()
        new_app = NewMagazynApp(new_window)

class NewMagazynApp:
    def __init__(self, master):
        # Dodaj atrybuty df_excel i df_subiekt
        self.df_excel = None
        self.df_subiekt = None
        self.master = master
        self.master.title("Magazyn")
        #self.master.iconbitmap("logo_ico.ico")
        # Ustawienie okna na środku ekranu
        window_width = 250
        window_height = 250
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        master.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Dodaj dowolne elementy do nowego okna
        label_quit = ttk.Label(self.master, text="Dostępne opcje", font=("TkDefaultFont", 10, "bold"))
        label_quit.grid(row=0, column=0, pady=10, padx=10)
        # Dolna stopka
        label_opcje = ttk.Label(self.master, text="© 2023 Dakro Bosch Service Autor: Mike Boro", font=("Arial", 8))
        label_opcje.grid(row=5, column=0, pady=10, padx=10)

        # Zmień status produktu
        opcja1 = ttk.Button(self.master, text="POMPY CR", command=self.porownanie)
        opcja1.grid(row=1, column=0, pady=10)

        # Zmień status produktu
        opcja2 = ttk.Button(self.master, text="TURBOSPRĘŻARKI", command=self.uruchom_porownanie_turbo)
        #opcja2.configure(state="disabled")
        opcja2.grid(row=2, column=0, pady=10)

        # Zmień status produktu
        opcja3 = ttk.Button(self.master, text="WTRYSKIWACZE / POMPOWTRYSKI", command=self.uruchom_porownanie_wtryskiwacze)
        #opcja3.configure(state="disabled")
        opcja3.grid(row=3, column=0, pady=10)

        # Zmień status produktu
        #opcja4 = ttk.Button(self.master, text="POMPOWTRYSKI", command=self.uruchom_porownanie_pompowtryski)
        #opcja4.configure(state="disabled")
        #opcja4.grid(row=4, column=0, pady=10)
    def porownanie(self):
        # Tutaj dodano nowe okno i zamknięcie poprzedniego
        self.open_porownanie()
    def open_porownanie(self):
        new_window = tk.Tk()
        new_window.title("Magazyn")
        # Ustawienie okna na środku ekranu
        window_width = 1350
        window_height = 500
        screen_width = new_window.winfo_screenwidth()
        screen_height = new_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        new_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        #new_window.iconbitmap("logo_ico.ico")
        # Ustawienia czcionki
        custom_font = font.nametofont("TkDefaultFont")
        custom_font.configure(size=custom_font.cget("size"))

        # Przycisk porównaj dane
        self.button_compare = tk.Button(new_window, text="Porównaj dane", command=self.uruchom_porownanie, font=custom_font)
        self.button_compare.pack(side=tk.TOP, pady=10, ipadx=10)

        # Pasek postępu
        self.progress = ttk.Progressbar(new_window, orient="horizontal", length=self.master.winfo_width() // 2, mode="determinate")
        self.progress.pack(side=tk.TOP, pady=20, fill=tk.X)

        # Lewa strona
        self.left_frame = tk.Frame(new_window)
        self.left_frame.pack(side=tk.LEFT, padx=10, pady=10, anchor="n")

        self.label_excel = tk.Label(self.left_frame, text="Dane Excel", font=custom_font)
        self.label_excel.grid(row=0, column=0, pady=(0, 10), sticky="w")

        # Zwiększono szerokość pola tekstowego
        self.entry_excel = tk.Entry(self.left_frame, state="readonly", font=custom_font, width=50)
        self.entry_excel.grid(row=1, column=0, pady=10)

        self.button_excel = tk.Button(self.left_frame, text="Wczytaj plik", command=self.wczytaj_excel, font=custom_font)
        self.button_excel.grid(row=2, column=0, pady=(10, 0), ipadx=10)

        self.tree_excel = ttk.Treeview(self.left_frame, columns=("Numer", "Numer Seryjny", "Ilość"), show="headings")
        self.tree_excel.heading("Numer", text="Numer")
        self.tree_excel.heading("Numer Seryjny", text="Numer Seryjny")
        self.tree_excel.heading("Ilość", text="Ilość")
        self.tree_excel.grid(row=3, column=0, pady=10)

        scrollbar_excel = tk.Scrollbar(self.left_frame, command=self.tree_excel.yview)
        scrollbar_excel.grid(row=3, column=1, pady=10, sticky="ns")
        self.tree_excel.config(yscrollcommand=scrollbar_excel.set)

        # Prawa strona
        self.right_frame = tk.Frame(new_window)
        self.right_frame.pack(side=tk.RIGHT, padx=10, pady=10, anchor="n")

        self.label_subiekt = tk.Label(self.right_frame, text="Dane Subiekt", font=custom_font)
        self.label_subiekt.grid(row=0, column=0, pady=(0, 10), sticky="w")

        # Zwiększono szerokość pola tekstowego
        self.entry_subiekt = tk.Entry(self.right_frame, state="readonly", font=custom_font, width=50)
        self.entry_subiekt.grid(row=1, column=0, pady=10)

        self.button_subiekt = tk.Button(self.right_frame, text="Wczytaj plik", command=self.wczytaj_subiekt, font=custom_font)
        self.button_subiekt.grid(row=2, column=0, pady=(10, 0), ipadx=10)

        self.tree_subiekt = ttk.Treeview(self.right_frame, columns=("Numer", "Numer Seryjny", "Ilość"), show="headings")
        self.tree_subiekt.heading("Numer", text="Numer")
        self.tree_subiekt.heading("Numer Seryjny", text="Numer Seryjny")
        self.tree_subiekt.heading("Ilość", text="Ilość")
        self.tree_subiekt.grid(row=3, column=0, pady=10)

        scrollbar_subiekt = tk.Scrollbar(self.right_frame, command=self.tree_subiekt.yview)
        scrollbar_subiekt.grid(row=3, column=1, pady=10, sticky="ns")
        self.tree_subiekt.config(yscrollcommand=scrollbar_subiekt.set)

        # Inicjalizacja danych różnicowych
        self.diff_data = []
    def uruchom_porownanie(self):
        # Uruchomienie aplikacji z innego pliku (porownanie.py)
        subprocess.run(["python", "porownanie.py"])
    def uruchom_porownanie_turbo(self):
        # Uruchomienie aplikacji z innego pliku (porownanie.py)
        subprocess.run(["python", "porownanie_turbo.py"])
    def uruchom_porownanie_wtryskiwacze(self):
        # Uruchomienie aplikacji z innego pliku (porownanie.py)
        subprocess.run(["python", "porownanie_wtryskiwacze.py"])
    def uruchom_porownanie_pompowtryski(self):
        # Uruchomienie aplikacji z innego pliku (porownanie.py)
        subprocess.run(["python", "porownanie_pompowtryski.py"])
    def wczytaj_excel(self):
        file_path_excel = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path_excel:
            self.entry_excel.configure(state="normal")
            self.entry_excel.delete(0, tk.END)
            self.entry_excel.insert(0, file_path_excel)
            self.entry_excel.configure(state="readonly")

            try:
                workbook = openpyxl.load_workbook(file_path_excel)
                sheet = workbook.active

                nazwa_pompy = []
                for row in sheet.iter_rows(min_row=2, max_col=5):
                    nazwa = str(row[1].value)
                    if row[4].value == "tak" and nazwa and nazwa.strip() and nazwa != "TYP PRODUKTU" and nazwa.lower() != "szukaj" and nazwa.lower() != "none":
                        nazwa_sformatowana = nazwa.lower().replace(" ", "")
                        nazwa_pompy.append(nazwa_sformatowana)

                nazwy_zliczone = Counter(nazwa_pompy)
                df_excel = pd.DataFrame({'Numer Seryjny': list(nazwy_zliczone.keys()), 'Ilość': list(nazwy_zliczone.values())})
                df_excel['Numer'] = range(1, len(df_excel) + 1)

                self.df_excel = df_excel

                self.fill_treeview(self.tree_excel, df_excel)

                # Save the loaded file
                self.save_loaded_file(file_path_excel, df_excel, 'Numer Seryjny', 'Ilość', 'Excel')

            except Exception as e:
                messagebox.showerror("Błąd wczytywania pliku", f"Wystąpił błąd podczas wczytywania pliku Excel: {e}")
                print("Błąd wczytywania pliku", f"Wystąpił błąd podczas wczytywania pliku Excel: {e}")

    def wczytaj_subiekt(self):
        file_path_subiekt = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path_subiekt:
            self.entry_subiekt.configure(state="normal")
            self.entry_subiekt.delete(0, tk.END)
            self.entry_subiekt.insert(0, file_path_subiekt)
            self.entry_subiekt.configure(state="readonly")

            try:
                df_subiekt = pd.read_excel(file_path_subiekt, usecols=[0, 1, 4], names=['Numer Seryjny', 'Nazwa', 'Ilość'])
                df_subiekt['Ilość'] = pd.to_numeric(df_subiekt['Ilość'], errors='coerce')
                df_subiekt = df_subiekt[df_subiekt['Ilość'] != 0]
                df_subiekt['Nazwa'] = df_subiekt['Nazwa'].apply(lambda x: x.split(' ', 1)[0])
                df_subiekt = df_subiekt[df_subiekt['Nazwa'].str.contains("POMPA", case=False)]

                self.usun_spacje(df_subiekt)
                self.usun_duze_litery(df_subiekt)
                df_subiekt['Numer'] = range(1, len(df_subiekt) + 1)

                self.fill_treeview(self.tree_subiekt, df_subiekt)

                # Save the loaded file
                self.save_loaded_file(file_path_subiekt, df_subiekt, 'Numer Seryjny', 'Ilość', 'Subiekt')

            except Exception as e:
                messagebox.showerror("Błąd wczytywania pliku", f"Wystąpił błąd podczas wczytywania pliku Excel: {e}")
                print("Błąd wczytywania pliku", f"Wystąpił błąd podczas wczytywania pliku Excel: {e}")

    def save_loaded_file(self, file_path, df, col1, col2, source):
        # Save only the specified columns
        selected_columns = [col1, col2]
        df_selected = df[selected_columns]

        # Determine the prefix for the file name based on the source
        prefix = ""
        if source.lower() == "excel":
            prefix = "Excel"
        elif source.lower() == "subiekt":
            prefix = "Subiekt"

        # Get the directory of the script
        script_dir = os.path.dirname(os.path.abspath(__file__))

        # Create the "Dane" subfolder if it doesn't exist
        dane_folder = os.path.join(script_dir, "Dane_Pompy")
        if not os.path.exists(dane_folder):
            os.makedirs(dane_folder)

        # Save the loaded file with the selected columns and prefix in the "Dane" folder
        save_path = os.path.join(dane_folder, os.path.basename(file_path).replace(".xlsx", f"-DANE.xlsx"))
        df_selected.to_excel(save_path, index=False, columns=selected_columns)
        
        messagebox.showinfo("Plik zapisany", f"Plik został pomyślnie zapisany jako {save_path}")
    
    def usun_spacje(self, df):
        # Usuń spacje z kolumny 'Numer Seryjny'
        df['Numer Seryjny'] = df['Numer Seryjny'].astype(str).str.replace(' ', '')
    def porownaj_dane(self):
        response = messagebox.askquestion("Potwierdzenie", "Czy na pewno chcesz kontynuować porównywanie danych?")
        if response == "yes":
            # Uruchom pasek postępu od lewej do prawej przez 3 sekundy
            self.progress.start(10)
            self.master.after(1000, self.zatrzymaj_postep)

            # Pobierz dane
            path_excel = self.entry_excel.get()
            path_subiekt = self.entry_subiekt.get()

            self.df_excel = pd.read_excel(path_excel, header=None, names=['Numer Seryjny', 'Ilość'])
            self.usun_spacje(self.df_excel)  
            self.usun_duze_litery(self.df_excel)  
            self.df_excel['Numer'] = range(1, len(self.df_excel) + 1)  

            self.df_subiekt = pd.read_excel(path_subiekt, usecols=[0, 4], names=['Numer Seryjny', 'Ilość'])
            self.usun_spacje(self.df_subiekt)  
            self.usun_duze_litery(self.df_subiekt)  
            self.df_subiekt = self.df_subiekt[self.df_subiekt['Ilość'] != 0]  
            self.df_subiekt['Numer'] = range(1, len(self.df_subiekt) + 1)  

            # Znajdź różnice
            self.znajdz_roznice()

        else:
            messagebox.showinfo("Anulowano", "Porównywanie danych zostało anulowane.")

    def znajdz_roznice(self):
        # Zresetuj dane różnicowe
        self.diff_data = []

        # Znajdź różnice w danych
        for index, row_excel in self.df_excel.iterrows():
            numer_seryjny_excel = row_excel['Numer Seryjny']
            ilosc_excel = row_excel['Ilość']

            # Sprawdź, czy numer seryjny istnieje w danych Subiekt
            row_subiekt = self.df_subiekt[self.df_subiekt['Numer Seryjny'] == numer_seryjny_excel]

            if not row_subiekt.empty:
                ilosc_subiekt = row_subiekt['Ilość'].values[0]

                # Porównaj ilość
                if ilosc_excel != ilosc_subiekt:
                    self.diff_data.append((numer_seryjny_excel, ilosc_excel, ilosc_subiekt))
            else:
                # Jeżeli numer seryjny nie istnieje w danych Subiekt, dodaj jako różnicę
                self.diff_data.append((numer_seryjny_excel, ilosc_excel, 0))

        # Sprawdź czy istnieją numery seryjne w danych Subiekt, które nie istnieją w danych Excel
        for index, row_subiekt in self.df_subiekt.iterrows():
            numer_seryjny_subiekt = row_subiekt['Numer Seryjny']

            # Sprawdź, czy numer seryjny istnieje w danych Excel
            row_excel = self.df_excel[self.df_excel['Numer Seryjny'] == numer_seryjny_subiekt]

            if row_excel.empty:
                # Jeżeli numer seryjny nie istnieje w danych Excel, dodaj jako różnicę
                self.diff_data.append((numer_seryjny_subiekt, 0, row_subiekt['Ilość']))



    def zatrzymaj_postep(self):
        # Zatrzymaj pasek postępu
        self.progress.stop()

        # Wyświetl powiadomienie o poprawnym porównaniu danych
        messagebox.showinfo("Powiadomienie", "Dane poprawnie porównane!")

        # Włącz przycisk "Pokaż różnice" po zakończeniu porównywania
        self.button_show_diff.configure(state="normal")

    def pokaz_roznice(self):
        if not self.diff_data:
            messagebox.showinfo("Brak różnic", "Nie znaleziono żadnych różnic.")
        else:
            # Utwórz nowe okno
            diff_window = tk.Toplevel()
            diff_window.title("Różnice")
            # Ustawienie okna na środku ekranu
            window_width = 600
            window_height = 600
            screen_width = diff_window.winfo_screenwidth()
            screen_height = diff_window.winfo_screenheight()
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            diff_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
            #diff_window.iconbitmap("logo_ico.ico")
            # Dodaj etykietę do nowego okna
            label = tk.Label(diff_window, text="Numery różniących się pozycji:", font=font.nametofont("TkDefaultFont"))
            label.pack()

            # Dodaj Treeview do nowego okna
            columns = ("Numer Seryjny", "Ilość - Excel", "Ilość - Subiekt")
            tree = ttk.Treeview(diff_window, columns=columns, show="headings")
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=100)

            tree.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

            # Dodaj suwak do nowego okna
            scrollbar = ttk.Scrollbar(diff_window, orient="vertical", command=tree.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            tree.configure(yscrollcommand=scrollbar.set)

            # Wypełnij Treeview danymi (bez spacji)
            for diff_item in self.diff_data:
                tree.insert("", "end", values=diff_item)

            # Dodaj przycisk eksportu
            export_button = tk.Button(diff_window, text="Eksportuj", command=self.eksportuj_roznice, font=font.nametofont("TkDefaultFont"))
            export_button.pack(side=tk.BOTTOM, pady=10)

    def usun_spacje_w_danych_roznicowych(self):
        # Usuń spacje z danych różnicowych
        for i in range(len(self.diff_data)):
            self.diff_data[i] = tuple(val.replace(" ", "") if isinstance(val, str) else val for val in self.diff_data[i])

    def eksportuj_roznice(self):
        # Wybierz lokalizację do zapisu pliku Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        if file_path:
            # Zastosuj usuwanie spacji w danych różnicowych
            self.usun_spacje_w_danych_roznicowych()

            # Zapisz różnice do pliku Excel
            if self.diff_data:
                # Jeśli istnieją różnice, utwórz ramkę danych
                df_diff = pd.DataFrame(self.diff_data, columns=["Numer Seryjny", "Ilość (Excel)", "Ilość (Subiekt)"])
            else:
                # Jeśli brak różnic, utwórz pustą ramkę danych
                df_diff = pd.DataFrame(columns=["Numer Seryjny", "Ilość (Excel)", "Ilość (Subiekt)"])

            df_diff.to_excel(file_path, index=False)

            messagebox.showinfo("Eksport zakończony", "Różnice zostały pomyślnie zapisane do pliku Excel.")

    def fill_treeview(self, treeview, df):
        # Wyczyść Treeview
        for child in treeview.get_children():
            treeview.delete(child)

        # Wypełnij Treeview danymi
        for index, row in df.iterrows():
            values = (row['Numer'], row['Numer Seryjny'], row['Ilość'])
            treeview.insert("", "end", values=values)
    def usun_duze_litery(self, df):
        # Usuń duże litery z kolumny 'Numer Seryjny'
        df['Numer Seryjny'] = df['Numer Seryjny'].apply(lambda x: str(x).lower())
        

def main():
    root = tk.Tk()
    app = MagazynApp(root)

    # Ustawienie okna na środku ekranu
    window_width = 300
    window_height = 250
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    #root.iconbitmap("logo_ico.ico")

    root.mainloop()

if __name__ == "__main__":
    main()