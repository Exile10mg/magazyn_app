
import tkinter as tk
from tkinter import filedialog
from tkinter import font
from tkinter import messagebox
import pandas as pd
from tkinter import ttk, simpledialog
import subprocess


class MagazynApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Magazyn")
        # Ustawienie okna na środku ekranu
        window_width = 1350
        window_height = 500
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.master.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Ustawienia czcionki
        custom_font = font.nametofont("TkDefaultFont")
        custom_font.configure(size=custom_font.cget("size"))

        # Przycisk porównaj dane
        self.button_compare = tk.Button(self.master, text="Porównaj dane", command=self.porownaj_dane, font=custom_font)
        self.button_compare.pack(side=tk.TOP, pady=10, ipadx=10)

        # Przycisk pokaż różnice
        self.button_show_diff = tk.Button(self.master, text="Pokaż różnice", command=self.pokaz_roznice, font=custom_font)
        self.button_show_diff.pack(side=tk.TOP, pady=10, ipadx=10)
        self.button_show_diff.configure(state="disabled")  # Domyślnie wyłączony

        # Pasek postępu
        self.progress = ttk.Progressbar(self.master, orient="horizontal", length=self.master.winfo_width() // 2, mode="determinate")
        self.progress.pack(side=tk.TOP, pady=20, fill=tk.X)

        # Lewa strona
        self.left_frame = tk.Frame(self.master)
        self.left_frame.pack(side=tk.LEFT, padx=10, pady=10, anchor="n")

        self.label_excel = tk.Label(self.left_frame, text="Dane Plik(1)", font=custom_font)
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
        self.right_frame = tk.Frame(self.master)
        self.right_frame.pack(side=tk.RIGHT, padx=10, pady=10, anchor="n")

        self.label_subiekt = tk.Label(self.right_frame, text="Dane Plik(2)", font=custom_font)
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
    def wczytaj_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.entry_excel.configure(state="normal")
            self.entry_excel.delete(0, tk.END)
            self.entry_excel.insert(0, file_path)
            self.entry_excel.configure(state="readonly")

            df = pd.read_excel(file_path, header=None, names=['Numer Seryjny', 'Ilość'])
            self.usun_spacje(df)  # Dodaj usuwanie spacji
            self.usun_duze_litery(df)  # Dodaj usuwanie dużych liter
            df['Numer'] = range(1, len(df) + 1)

            self.fill_treeview(self.tree_excel, df)

    def wczytaj_subiekt(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.entry_subiekt.configure(state="normal")
            self.entry_subiekt.delete(0, tk.END)
            self.entry_subiekt.insert(0, file_path)
            self.entry_subiekt.configure(state="readonly")

            df = pd.read_excel(file_path, header=None, names=['Numer Seryjny', 'Ilość'])
            self.usun_spacje(df)  # Dodaj usuwanie spacji
            self.usun_duze_litery(df)  # Dodaj usuwanie dużych liter
            df['Numer'] = range(1, len(df) + 1)

            self.fill_treeview(self.tree_subiekt, df)

    def porownaj_dane(self):
        response = messagebox.askquestion("Potwierdzenie", "Czy na pewno chcesz kontynuować porównywanie danych?")
        if response == "yes":
            # Uruchom pasek postępu od lewej do prawej przez 3 sekundy
            self.progress.start(10)
            self.master.after(1000, self.zatrzymaj_postep)

            # Pobierz dane z plików
            path_excel = self.entry_excel.get()
            path_subiekt = self.entry_subiekt.get()

            df_excel = pd.read_excel(path_excel, header=None, names=['Numer Seryjny', 'Ilość'])
            self.usun_spacje(df_excel)  # Dodaj usuwanie spacji
            self.usun_duze_litery(df_excel)  # Dodaj usuwanie dużych liter
            df_excel['Numer'] = range(1, len(df_excel) + 1)  # Dodaj kolumnę 'Numer' z numerami porządkowymi

            df_subiekt = pd.read_excel(path_subiekt, header=None, names=['Numer Seryjny', 'Ilość'])
            self.usun_spacje(df_subiekt)  # Dodaj usuwanie spacji
            self.usun_duze_litery(df_subiekt)  # Dodaj usuwanie dużych liter
            df_subiekt['Numer'] = range(1, len(df_subiekt) + 1)  # Dodaj kolumnę 'Numer' z numerami porządkowymi

            # Porównaj dane i podświetl różnice
            self.podswietl_roznice(df_excel, df_subiekt)
        else:
            messagebox.showinfo("Anulowano", "Porównywanie danych zostało anulowane.")

    def usun_spacje(self, df):
        # Usuń spacje z kolumny 'Numer Seryjny'
        df['Numer Seryjny'] = df['Numer Seryjny'].astype(str).str.replace(' ', '')

    def zatrzymaj_postep(self):
        # Zatrzymaj pasek postępu
        self.progress.stop()

        # Wyświetl powiadomienie o poprawnym porównaniu danych
        messagebox.showinfo("Powiadomienie", "Dane poprawnie porównane!")

        # Włącz przycisk "Pokaż różnice" po zakończeniu porównywania
        self.button_show_diff.configure(state="normal")

    def podswietl_roznice(self, df1, df2):
        self.diff_data = []  # Zresetuj dane różnicowe

        for index, row in df1.iterrows():
            nr_seryjny = row['Numer Seryjny']
            ilosc_excel = row['Ilość']

            # Sprawdź, czy numer seryjny występuje w obu ramkach danych
            if nr_seryjny in df2['Numer Seryjny'].values:
                ilosc_subiekt = df2[df2['Numer Seryjny'] == nr_seryjny]['Ilość'].values[0]
                lp_subiekt = df2[df2['Numer Seryjny'] == nr_seryjny]['Numer'].values[0]
            else:
                ilosc_subiekt = 0
                lp_subiekt = None  # Użyj None, aby wskazać, że numer seryjny nie występuje w df2

            if ilosc_excel != ilosc_subiekt:
                self.podswietl_wiersz(self.tree_excel, row['Numer'])

                # Podświetl wiersz w widżecie tekstowym Subiekt, jeśli numer seryjny istnieje
                if lp_subiekt is not None:
                    self.podswietl_wiersz(self.tree_subiekt, lp_subiekt)

                # Dodaj dane różnicowe do listy
                self.diff_data.append((nr_seryjny, ilosc_excel, ilosc_subiekt))
            elif lp_subiekt is None:
                # Jeśli numer seryjny nie występuje w Subiekcie, dodaj komunikat do listy różnic
                self.diff_data.append((nr_seryjny, ilosc_excel, "---"))

        # Podświetl wiersze w Subiekcie, gdzie numery seryjne występują, ale nie w Excelu
        for index, row in df2.iterrows():
            nr_seryjny = row['Numer Seryjny']
            if nr_seryjny not in df1['Numer Seryjny'].values:
                lp_subiekt = row['Numer']
                self.podswietl_wiersz(self.tree_subiekt, lp_subiekt)
                self.diff_data.append((nr_seryjny, "---", row['Ilość']))

    def podswietl_wiersz(self, treeview, line_number):
        item_id = treeview.get_children()[line_number-1]  # Identyfikator wiersza w Treeview
        treeview.item(item_id, tags=('highlight',))
        treeview.see(item_id)  # Przewiń do podświetlonego wiersza

    def pokaz_roznice(self):
        if not self.diff_data:
            messagebox.showinfo("Brak różnic", "Nie znaleziono żadnych różnic.")
        else:
            # Utwórz nowe okno
            diff_window = tk.Toplevel(self.master)
            diff_window.title("Różnice")
            # Ustawienie okna na środku ekranu
            window_width = 800
            window_height = 600
            screen_width = diff_window.winfo_screenwidth()
            screen_height = diff_window.winfo_screenheight()
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            diff_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

            # Dodaj etykietę do nowego okna
            label = tk.Label(diff_window, text="Numery różniących się pozycji:", font=font.nametofont("TkDefaultFont"))
            label.pack()

            # Dodaj Treeview do nowego okna
            columns = ("Numer Seryjny", "Ilość - Plik(1)", "Ilość - Plik(2)")
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
                df_diff = pd.DataFrame(self.diff_data, columns=["Numer Seryjny", "Ilość Plik(1)", "Ilość Plik(2)"])
            else:
                # Jeśli brak różnic, utwórz pustą ramkę danych
                df_diff = pd.DataFrame(columns=["Numer Seryjny", "Ilość Plik(1)", "Ilość Plik(2)"])

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
    root.mainloop()

if __name__ == "__main__":
    main()