import pandas as pd
from datetime import datetime
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox

def show_notification(message):
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo('Powiadomienie', message)

def send_email_via_outlook(subject, body, recipient):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        print("Wiadomość e-mail została wysłana.")
        show_notification("Wiadomość e-mail została wysłana.")
    except Exception as e:
        print(f"Wystąpił błąd przy wysyłaniu e-maila: {e}")
        show_notification(f"Wystąpił błąd przy wysyłaniu e-maila: {e}")

def read_email_from_file():
    with open('email.txt', 'r') as file:
        return file.readline().strip()

def process_date(date_column, modulo, today, column_name, email_body, notification_body):
    try:
        if pd.notna(modulo) and pd.notna(date_column):
            if isinstance(date_column, datetime):
                date_column = date_column.date()
            elif isinstance(date_column, str):
                date_column = datetime.strptime(date_column.strip(), '%d.%m.%Y').date()

            days_to_date = (date_column - today).days
            if days_to_date == 30:
                message = f"Modulo {modulo}: Pozostało 30 dni do {column_name}.\n"
                email_body += message
                notification_body += message
            elif days_to_date <= 7 and days_to_date >= 0:
                message = f"Modulo {modulo}: Pozostało {days_to_date} dni do {column_name}.\n"  
                email_body += message 
                notification_body += message
            elif days_to_date < 0:
                message = f"Modulo {modulo}: Termin na {column_name} upłynął. Minęło {abs(days_to_date)} dni.\n"
                email_body += message
                notification_body += message
    except ValueError as e:
        error_message = f"Błąd przy przetwarzaniu daty '{date_column}' dla modulo {modulo}. Szczegóły błędu: {e}"
        print(error_message)
        notification_body += error_message + "\n"
    return email_body, notification_body

def notify_about_policies():
    recipient_email = read_email_from_file()

    try: 
        df = pd.read_excel('.xlsx')
        today = datetime.now().date()

        email_body = ""
        notification_body = ""

        for _, row in df.iterrows():
            modulo = row['MODULO']
            policy_date = row['data polisy']
            email_body, notification_body = process_date(policy_date, modulo, today, "data polisy", email_body, notification_body)
            fulfillment_date = row['DATA SPEŁNIENIA']
            email_body, notification_body = process_date(fulfillment_date, modulo, today, "DATA SPEŁNIENIA", email_body, notification_body)

        if notification_body:  
            show_notification(notification_body)
            send_email_via_outlook("Powiadomienia", email_body, recipient_email)
        else:
            message = "Wszystkie daty są aktualne."
            print(message)
            show_notification(message) 

    except Exception as e:
        print(f"Wystąpił błąd: {e}")
        show_notification(f"Wystąpił błąd: {e}")


notify_about_policies()