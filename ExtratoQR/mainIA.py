from treatment import treatment_file
from tkinter import filedialog
from tkinter import messagebox



try:
    extrato = filedialog.askopenfilename(filetypes=[("Extrato QR", "*.xlsx")])
    reltech = filedialog.askopenfilename(filetypes=[("Relatório Tech", "*.xlsx")])
    dataframe = treatment_file.TreatmentDataFrame(extrato, reltech)
    dataframe.pagments_more_than_one()
    messagebox.showinfo(title='ExtratoQR', message='Conferencia realizada com sucesso!')
except Exception as e:
    messagebox.showerror(title='ExtratoQR', message=e)


