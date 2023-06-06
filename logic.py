import xlrd
import smtplib
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# create tkinter window to select the file
def select_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        process_file(file_path)

# process selected file
def process_file(file_path):
    xls_file = xlrd.open_workbook(file_path) # reads excel sheet
    xl = xls_file.sheet_by_index(0) # ignores the 1st row i.e., name, id,....
    l = []
    m = []
    for i in range(1,xl.nrows):
        l.append(xl.row_values(i))
    n = 7.5 # GPA cutoff
    for i in ((l)):
        for j in i[2:]:
            if int(j) < n:
                m.append(i[1])
                break

    sender_mail = ""
    password = ""
    receiver_mail = ""
    message = """
        Dear student,
            This is to inform you that you got less GPA in semester exam.
    """
    try:
        server_connect = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server_connect.login(sender_mail, password)
        for i in m:
            receiver_mail = i
            server_connect.sendmail(sender_mail, receiver_mail, message)
        server_connect.quit()
        messagebox.showinfo("Success", "Mails have been sent!")
    except Exception as e:
        messagebox.showerror("Error", f"Error sending mails:\n{e}")

# create main window
root = tk.Tk()
root.title("GPA Report")

# create file selection button
file_button = tk.Button(root, text="Select File", command=select_file)
file_button.pack(pady=10)

# run main loop
root.mainloop()
