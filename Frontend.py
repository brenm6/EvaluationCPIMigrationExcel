import tkinter as tk
from tkinter import filedialog, messagebox

class ExcelFrontend:
    def __init__(self, master):
        self.master = master
        master.title("Excel Evaluation Tool")
        master.configure(bg="black")
        master.geometry("500x250")  # Fenster größer machen

        font_big_bold = ("Arial", 14, "bold")
        font_status = ("Arial", 12, "bold")

        self.label = tk.Label(
            master,
            text="Bitte wählen Sie ein CPI Evaluation Excel aus dem  ('evaluation_run_results_PA3_...') aus:",
            font=font_big_bold,
            fg="white",
            bg="black",
            wraplength=480,  # Textumbruch passend zur Fensterbreite
            justify="center"
        )
        self.label.pack(pady=20)

        self.upload_button = tk.Button(
            master,
            text="Datei auswählen",
            command=self.upload_file,
            font=font_big_bold,
            fg="white",
            bg="#0078D7",  # Microsoft-Blau
            activebackground="#005A9E",
            activeforeground="white",
            width=20,
            height=2,
            bd=0,
            highlightthickness=0,
            cursor="hand2"
        )
        self.upload_button.pack(pady=10)

        self.status_label = tk.Label(
            master,
            text="",
            font=font_status,
            fg="white",
            bg="black"
        )
        self.status_label.pack(pady=10)

    def upload_file(self):
        from Excel_Manager import ExcelManager  # Import hier lokal!
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
            title="Excel-Datei auswählen"
        )
        if file_path:
            self.status_label.config(text="Datei wird verarbeitet...dauert je nach Dateigröße einige Sekunden")
            self.master.update()
            try:
                excel_manager = ExcelManager(file_path)
                sheet_to_add = excel_manager.create_sheet('Evaluation', 2)
                excel_manager.set_columns(sheet_to_add)
                excel_manager.fill_sheet(sheet_to_add)
                
                # Entferne alle anderen Sheets außer "Evaluation"
                for ws in list(excel_manager.workbook.sheetnames):
                    if ws != "Evaluation":
                        std = excel_manager.workbook[ws]
                        excel_manager.workbook.remove(std)
                
                excel_manager.save()
                self.status_label.config(text="Verarbeitung abgeschlossen und gespeichert!")
                messagebox.showinfo("Fertig", "Die Datei wurde erfolgreich verarbeitet und gespeichert.")
            except Exception as e:
                self.status_label.config(text="Fehler bei der Verarbeitung!")
                if "Permission denied" in str(e):
                    messagebox.showerror(
                        "Fehler",
                        "Die Datei konnte nicht gespeichert werden.\nBitte schließen Sie die geöffnete Excel-Datei, und versuchen Sie es erneut."
                    )
                else:
                    messagebox.showerror("Fehler", f"Es ist ein Fehler aufgetreten:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelFrontend(root)
    root.mainloop()