import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk
from classes import MyApi
import sys
import io
import threading
import ssl


class App:
    def __init__(self):
        
        # Tentative de bypass des vérifications ssl (fontionnement non garanti)
        ssl._create_default_https_context = ssl._create_unverified_context
        self.api = MyApi()
        self.root = tk.Tk()
        self.root.title("ETS OpenAlex Analyzer")
        self.is_processing = False
        self.progress_window = None

        # Configuration de la grille
        self.root.grid_rowconfigure(3, weight=1)
        self.root.grid_columnconfigure((0, 1, 2, 3), weight=1)
        
        # Validation numérique pour les années
        val_num = self.root.register(self.__validate_year_input)
        
        # Widgets pour les années
        tk.Label(self.root, text="Année de début:").grid(row=0, column=0, sticky="w")
        self.start_entry = tk.Entry(self.root, width=8, validate="key", validatecommand=(val_num, '%P'))
        self.start_entry.insert(0, "2019")
        self.start_entry.grid(row=0, column=1, sticky="ew")
        
        tk.Label(self.root, text="Année de fin:").grid(row=0, column=2, sticky="w")
        self.end_entry = tk.Entry(self.root, width=8, validate="key", validatecommand=(val_num, '%P'))
        self.end_entry.insert(0, "2023")
        self.end_entry.grid(row=0, column=3, sticky="ew")
        
        # Widgets pour le ROR collaborateur
        tk.Label(self.root, text="ROR Collaborateur:").grid(row=1, column=0, sticky="w")
        self.ror_entry = tk.Entry(self.root)
        self.ror_entry.insert(0, "https://ror.org/02feahw73") # ROR du CNRS par défaut
        self.ror_entry.grid(row=1, column=1, columnspan=3, sticky="ew")
        
        # Boutons
        buttons = [
            ("Récupérer les publications", self.fetch_works),
            ("Lister pays collaborateurs", self.show_collaborators),
            ("Générer rapport word", self.generate_report),
            ("Sujets principaux avec le collaborateur", self.analyze_collaboration)
        ]
        
        for i, (text, cmd) in enumerate(buttons):
            tk.Button(self.root, text=text, command=cmd).grid(
                row=2, column=i, sticky="ew", padx=2, pady=2
            )
        
        # Zone de logs
        self.log_area = scrolledtext.ScrolledText(self.root, state="disabled")
        self.log_area.grid(row=3, column=0, columnspan=4, sticky="nsew")
        
        # Redirection de la sortie
        self.output = io.StringIO()
        sys.stdout = self.output
        
        # Mise à jour périodique des logs
        self.update_log()

    def __validate_year_input(self, value):
        """Valide que l'entrée est numérique et <= 4 caractères"""
        if value == "" or (value.isdigit() and len(value) <= 4):
            return True
        return False

    def __get_validated_years(self):
        """Récupère et valide les années avec messages d'erreur"""
        try:
            start = int(self.start_entry.get())
            end = int(self.end_entry.get()) if self.end_entry.get() else start
            
            # Validation de la plage
            if not (1981 <= start <= 2025):
                raise ValueError("L'année de début doit être entre 1981 et 2025")
                
            if not (1981 <= end <= 2025):
                raise ValueError("L'année de fin doit être entre 1981 et 2025")
                
            if start > end:
                raise ValueError("L'année de début ne peut pas être supérieure à l'année de fin")
            
            return (start, end)
            
        except ValueError as e:
            messagebox.showerror("Erreur de saisie", str(e))
            return None

    def __show_processing(self):
        """Affiche une fenêtre de progression centrée"""
        self.is_processing = True
        self.should_stop = False
        
        # Création de la fenêtre
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("Traitement en cours")
        
        # Création du contenu
        frame = tk.Frame(self.progress_window)
        frame.pack(pady=10, padx=20)
        
        tk.Label(frame, text="Veuillez patienter, traitement en cours...").grid(row=0, column=0)
        self.progress = ttk.Progressbar(frame, mode='indeterminate')
        self.progress.grid(row=1, column=0, pady=5)
        self.progress.start()
        
        cancel_btn = tk.Button(frame, text="Annuler", command=self.__cancel_operation)
        cancel_btn.grid(row=2, column=0, pady=10)
        
        # Calcul du centrage
        self.progress_window.update_idletasks()  # Force le calcul des dimensions
        
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()
        
        progress_width = self.progress_window.winfo_width()
        progress_height = self.progress_window.winfo_height()
        
        x = main_x + (main_width - progress_width) // 2
        y = main_y + (main_height - progress_height) // 2
        
        self.progress_window.geometry(f"+{x}+{y}")  # Positionnement au centre
        
        # Configuration finale
        self.progress_window.transient(self.root)  # Lien hiérarchique
        self.progress_window.grab_set()
        self.progress_window.protocol("WM_DELETE_WINDOW", self.__cancel_operation)

    def __cancel_operation(self):
        """Gère la demande d'annulation par l'utilisateur"""
        self.should_stop = True
        self.__hide_processing()
        print("⏹ Traitement annulé par l'utilisateur")
        
    def __hide_processing(self):
        """Cache la fenêtre de progression"""
        if self.progress_window:
            self.progress_window.destroy()
        self.is_processing = False
        
    def __thread_wrapper(self, target_method, args=()):
        """Lance une méthode dans un thread séparé"""
        if self.is_processing:
            return
            
        self.should_stop = False  # Réinitialiser l'état d'annulation
        
        def thread_target():
            try:
                self.__show_processing()
                target_method(*args)
            finally:
                self.root.after(0, self.__hide_processing)
        
        threading.Thread(target=thread_target, daemon=True).start()

    def update_log(self):
        """Gestion des messages sur la fenêtre principale"""
        self.log_area.config(state="normal")
        self.log_area.insert("end", self.output.getvalue())
        self.log_area.see("end")
        self.log_area.config(state="disabled")
        self.output.truncate(0)
        self.output.seek(0)
        self.root.after(100, self.update_log)
        
    def fetch_works(self):
        years = self.__get_validated_years()
        if not years:
            return
        self.__thread_wrapper(self.__fetch_works_task, years)
    
    def __fetch_works_task(self, start, end):
        try:
            self.api.show_works(start, end, self._check_stop)
            if self.should_stop: return
            print(f"✅ Publications {start}-{end} récupérées avec succès")
        except Exception as e:
            if not self.should_stop:  # Ne pas afficher l'erreur si annulation
                messagebox.showerror("Erreur", str(e))
    
    def _check_stop(self):
        """Vérification des interruptions pendant les opérations"""
        if self.should_stop:
            raise Exception("Traitement interrompu par l'utilisateur")
        
    def show_collaborators(self):
        years = self.__get_validated_years()
        if not years:
            return
        self.__thread_wrapper(self.__show_collaborators_task, years)
    
    def __show_collaborators_task(self, start, end):
        self.api.show_collaborators(start, end, self._check_stop)
        print(f"✅ Liste des collaborateurs {start}-{end} générée")
    
    def generate_report(self):
        years = self.__get_validated_years()
        if not years:
            return
        self.__thread_wrapper(self.__generate_report_task, years)
    
    def __generate_report_task(self, start, end):
        self.api.generate_country_report(start, end, self._check_stop)
    
    def analyze_collaboration(self):
        try:
            ror = self.ror_entry.get().strip()
            if not ror:
                messagebox.showerror("Erreur", "ROR collaborateur requis")
                return
                
            years = self.__get_validated_years()
            if not years:
                return
                
            self.__thread_wrapper(self.__analyze_collaboration_task, (ror, *years))
            
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
    
    def __analyze_collaboration_task(self, ror, start, end):
        try:
            self.api.show_works_with_collaboration(ror, start, end, self._check_stop)
        except Exception as e :
            messagebox.showerror("Erreur", str(e))
            return
        
        print(f"✅ Analyse des collaborations {start}-{end} avec {ror} terminée")
    
    def on_close(self):
        sys.stdout = sys.__stdout__
        self.root.destroy()

if __name__ == "__main__":
    app = App()
    app.root.protocol("WM_DELETE_WINDOW", app.on_close)
    app.root.mainloop()