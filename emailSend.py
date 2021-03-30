from tkinter import *
from tkinter import messagebox, filedialog
import yagmail
from plyer import notification
import csv


class Sender:
    def __init__(self, root):
        self.root = root
        self.attach = []
        self.to_email_list = []
        self.user_email = ''
        self.user_password = ''

        label = Label(self.root, bg='#0099ff', bd=5, relief=FLAT)
        label.pack(fill=BOTH)

        self.main_frame = Frame(self.root, width=600, bd=5, relief=GROOVE)
        self.main_frame.pack(side=LEFT, fil=Y)

        self.to_label = Label(self.main_frame, text="To", font=('arial', 12, 'bold'))
        self.to_label.place(x=0, y=10)

        self.to_entry = Entry(self.main_frame, font=('arial', 12), width=54, bd=2)
        self.to_entry.place(x=50, y=10)

        self.csv_attach_btn = Button(self.main_frame, text='CSV', command=self.select_csv)
        self.csv_attach_btn.place(x=545, y=10)
        self.csv_attach_btn.config(state='disabled')

        self.subject_label = Label(self.main_frame, text="Subject", font=('arial', 12, 'bold'))
        self.subject_label.place(x=0, y=40)

        self.subject_entry = Entry(self.main_frame, font=('arial', 12), width=50, bd=2)
        self.subject_entry.place(x=74, y=40)

        self.text_frame = Frame(self.main_frame, bd=2, bg='sky blue')
        self.text_frame.place(x=0, y=80, height=450, width=585)

        scroll_bar = Scrollbar(self.text_frame)
        scroll_bar.pack(side=RIGHT, fill=Y)

        self.textbox = Text(self.text_frame, height=440, font=('arial', 12), yscrollcommand=scroll_bar.set)
        self.textbox.pack(fill=Y)
        scroll_bar.config(command=self.textbox.yview)

        self.btn_frame = Frame(self.main_frame, bd=2, bg='#0099ff')
        self.btn_frame.place(x=0, y=530, height=30, width=585)

        self.clear_btn = Button(self.btn_frame, text='Clear', command=lambda: self.textbox.delete(0.0, END))
        self.clear_btn.grid(row=0, column=0, ipadx=5, padx=5, sticky='nesw')
        self.clear_btn.config(state='disabled')

        # self.exit_btn = Button(self.btn_frame, text='Exit', command=self.root.destroy)
        # self.exit_btn.grid(row=0, column=1, ipadx=5, padx=5)

        self.attach_btn = Button(self.btn_frame, text='Attach', command=self.add_attach)
        self.attach_btn.grid(row=0, column=2, ipadx=5, padx=5, sticky='nesw')
        self.attach_btn.config(state='disabled')

        self.login_btn = Button(self.btn_frame, text='Login', command=self.log_in)
        self.login_btn.grid(row=0, column=3, ipadx=5, padx=5, sticky='nesw')

        self.send_btn = Button(self.btn_frame, text='Send', command=self.send_mail)
        self.send_btn.grid(row=0, column=4, ipadx=5, padx=5, sticky='nesw')
        self.send_btn.config(state='disabled')

    def log_in(self):
        self.top = Toplevel(self.root)
        self.password_entry = Entry(self.top, width=20, font=('arial', 12), bd=2)
        self.email_entry = Entry(self.top, width=20, font=('arial', 12), bd=2)
        self.sender_email = self.email_entry.get()
        self.top.title('Log in screen')
        self.top.geometry('400x150+600+200')
        self.top.configure(bg='sky blue')

        Label(self.top, text='Email:', font=('arial', 12, 'bold'), width=10).grid(row=0, column=0,
                                                                                               padx=20, pady=5)
        Label(self.top, text='Password:', font=('arial', 12, 'bold'), width=10).grid(row=1,
                                                                                                      column=0,
                                                                                                      padx=20,
                                                                                                      pady=5)

        self.email_entry.grid(row=0, column=1, padx=20, pady=20)

        self.password_entry.grid(row=1, padx=20, column=1, pady=5)
        self.password_entry.config(show='*')

        log_in = Button(self.top, text="Log in", width=10, font=('arial', 12, 'bold'), command=self.log_in_successful)
        log_in.grid(row=2, padx=30, column=0, pady=5)

        cancel_btn = Button(self.top, text="Cancel", width=10, font=('arial', 12, 'bold'), command=self.top.quit)
        cancel_btn.grid(row=2, padx=30, column=1, pady=5)

    def log_in_successful(self):
        self.sender_password = self.password_entry.get()

        if self.sender_email != '' or self.sender_password != '':
            self.user_email = self.email_entry.get()
            self.user_password = self.password_entry.get()

            messagebox.showinfo('Sender Says', 'Login details are saved for mail sending purpose')
            self.send_btn.config(state='normal')
            self.top.destroy()

        else:
            messagebox.showerror('Sender Says', 'Enter all fields correctly')
            self.top.destroy()

    def add_attach(self):
        files = filedialog.askopenfilenames()
        for file in files:
            self.attach.append(file)

    def send_mail(self):
        if self.to_entry.get() == '' or self.textbox.get(0.0, END) == '':
            messagebox.showwarning('Sender Says', "Fill all required fields correctly.")

        else:
            yag = yagmail.SMTP(user=self.user_email, password=self.user_password)

            if self.attach:
                for email in self.to_email_list:
                    yag.send(email, self.subject_entry.get(), self.textbox.get(0.0, END), self.attach)
                    self.attach = []
            else:
                for email in self.to_email_list:
                    yag.send(email, self.subject_entry.get(), self.textbox.get(0.0, END))

            notification.notify(title="Success", message="mail is sent", timeout=5)

    def select_csv(self):
        Tk().withdraw()
        text_file_extensions = ['*.csv']
        ftypes = [
            ('CSV Files', text_file_extensions),
            ('All files', '*'),
        ]
        csv_file = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=ftypes)

        with open(csv_file, 'rt')as file:
            data = csv.reader(file)
            for row in data:
                self.to_email_list.append(row)

        emails = ','.join([str(element) for element in self.to_email_list])
        emails = emails.replace("'", '').replace('[', '').replace(']', '')

        self.to_entry.insert(0, string=emails)


if __name__ == "__main__":
    root = Tk()
    Sender(root)
    root.title("Email Sender")
    root.geometry('590x600')
    root.resizable(0, 0)
    root.mainloop()
