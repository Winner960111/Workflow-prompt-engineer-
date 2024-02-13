from typing import Optional
from myfunction import *
from flask import Flask, render_template, request, url_for, jsonify, make_response
from flask_cors import CORS
import requests
import os
import sqlite3
import pickle
import asyncio

from langchain.chains.openai_functions import (
    create_openai_fn_chain,
    create_openai_fn_runnable,
    create_structured_output_chain,
    create_structured_output_runnable,
)
from langchain.chat_models import ChatOpenAI
from langchain.prompts import ChatPromptTemplate
from langchain.document_loaders import PyPDFLoader
from langchain.pydantic_v1 import BaseModel, Field
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain

from openai import OpenAI

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from base64 import urlsafe_b64decode, urlsafe_b64encode

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
import logging
from mimetypes import guess_type as guess_mime_type

import uuid
import time
import json
import re
from twilio.rest import Client
from dotenv import load_dotenv
load_dotenv()
app = Flask(__name__)
CORS(app, origins='*')
openai_api_key = os.environ.get("OPENAI_API_KEY")
SCOPES = ['https://mail.google.com/']
our_email = 'harryporter319193@gmail.com'

twilio_account_sid = os.environ.get("TWILIO_ACCOUNT_SID")
twilio_auth_token = os.environ.get("TWILIO_AUTH_TOKEN")
twilio_phone_number = os.environ.get("TWILIO_PHONE_NUMBER")
sms_phone_number = os.environ.get("SMS_NUMBER")


# twilio_client = Client(twilio_account_sid, twilio_auth_token)
openai_client = OpenAI(api_key=openai_api_key)

states = {}
timeslot = {}
with open('log.txt', 'w') as file:
    # Clear the contents of the file
    file.truncate(0)


def phone_message(message, to_number, option):  # to_number have to include + at first

    if option == "sms":
        from_ = f'{sms_phone_number}'
        to = f'{to_number}'
    else:
        from_ = f'whatsapp:{twilio_phone_number}'
        to = f'whatsapp:{to_number}'

    # twilio_client.messages.create(
    #     from_=from_,
    #     body=message,
    #     to=to
    # )


def gmail_authenticate():
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.with_subject(
                "harryporter391193@gmail.com")
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)


service = gmail_authenticate()


def build_email(destination, obj, body, attachments=[]):
    message = MIMEText(body)
    message['to'] = destination
    message['from'] = our_email
    message['subject'] = obj
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}

def send_email(destination, body):
    try:
        # Send the email...
        send_response = service.users().messages().send(userId="me", body=build_email(
            destination, "Message for Candidate", body)).execute()

        # Log the successful response
        print(f"Email sent; Message ID: {send_response['id']}")
    except Exception as e:
        print("An error occurred while sending an email.")
    filename = 'log.txt'

    # Open the file in append mode
    with open(filename, 'a') as file:
        # Append the new content to the file
        file.write(f"To candidate => {body}\n")


def extract_info(filename):
    loader = PyPDFLoader(f"./uploads/{filename}")
    input = loader.load()[0].page_content
    runnable = create_structured_output_runnable(Person, llm, extract_prompt)
    res = runnable.invoke({"input": input})

    info = {}
    info["name"] = res.name
    info["email"] = res.email
    info["number"] = res.number
    return info


def email_mark_as_read(id):
    service.users().messages().batchModify(userId='me', body={
        'ids': [id], 'removeLabelIds': ['UNREAD']}).execute()


def clean_msg(message):
    pattern = re.compile(
        r"On [A-Za-z]{3}, \w+ \d{1,2}, \d{4} at \d{1,2}:\d{2} (AM|PM)")
    clean_snippet = re.split(pattern, message)[0].strip()
    return clean_snippet
def set_step(email, step):
    conn = sqlite3.connect("mydb.sqlite")
    cur = conn.cursor()
    cur.execute("UPDATE resume SET step = ? WHERE email = ?",
                (step, email))
    conn.commit()
    cur.close()
    conn.close()

# send  simple email

def read_simple_email():
    global states, timeslot
    email_items = []
    # get the emails from SQLite DB
    with sqlite3.connect('mydb.sqlite') as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT email FROM resume")
        email_items = [row[0] for row in cursor.fetchall()]
    cursor.close()
    conn.close()

    for from_email in email_items:

        results = service.users().messages().list(
            userId='me', q=f"from:{from_email}").execute()
        email_messages = results.get('messages', [])

        with sqlite3.connect('mydb.sqlite') as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT step FROM resume WHERE email = ?", (from_email,))
            step = cursor.fetchone()[0]
        cursor.close()
        conn.close()

        print(f"On {from_email} current step is =>{step}")
        for message in email_messages:
            msg = service.users().messages().get(
                userId='me', id=message['id']).execute()

            if 'UNREAD' in msg['labelIds']:
                filename = 'log.txt'

                # Open the file in append mode
                with open(filename, 'a') as file:
                    # Append the new content to the file
                    file.write(
                        f"From candidate => {clean_msg(msg['snippet'])}\n")

                if step == 1:
                    res = init_answer(from_email, clean_msg(msg['snippet']))
                    print(res)
                    if (res == 'Yes'):
                        set_step(from_email, 2)
                        res = show_JD(from_email)
                        print(res)
                        send_email(from_email, res)
                    elif (res == 'No'):
                        res = reason_message()
                        print(res)
                        send_email(from_email, res)
                        set_step(from_email, 0)
                    else:
                        res = init_message(from_email, True)
                        print(res)
                        send_email(from_email, res)
                elif step == 0:
                    res = (clean_msg(msg['snippet']))
                    print(res)
                    send_email(from_email, end_message())
                    set_step(from_email, 10)
                elif step == 2:
                    res = show_JD_answer(from_email, clean_msg(msg['snippet']))
                    print(res)
                    if res == "Yes":
                        set_step(from_email, 3)
                        timeslot[from_email] = calendar_show()
                        print(timeslot[from_email])
                        send_email(from_email, timeslot[from_email])
                    elif res == "No":
                        call_res = end_message()
                        print(call_res)
                        send_email(from_email, call_res)
                    else:
                        call_res = show_JD(from_email)
                        print(call_res)
                        send_email(from_email, call_res)
                    res = JD_recruiter_answer(
                        from_email, clean_msg(msg['snippet']))
                    print(res)
                elif step == 3:
                    res = calendar_book(timeslot[from_email], clean_msg(
                        msg['snippet']), from_email)
                    print(res)
                    if res == "None":
                        timeslot[from_email] = calendar_show()
                        print(timeslot[from_email])
                        send_email(from_email, timeslot[from_email])
                    else:
                        call_res = reserve_message(from_email, res)
                        set_step(from_email, 10)
                        print(call_res)
                        send_email(from_email, call_res)
                   

                email_mark_as_read(message['id'])

# send medium email
def read_md_email():
    global states, timeslot
    email_items = []
    # get the emails from SQLite DB
    with sqlite3.connect('mydb.sqlite') as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT email FROM resume")
        email_items = [row[0] for row in cursor.fetchall()]
    cursor.close()
    conn.close()

    for from_email in email_items:

        results = service.users().messages().list(
            userId='me', q=f"from:{from_email}").execute()
        email_messages = results.get('messages', [])

        with sqlite3.connect('mydb.sqlite') as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT step FROM resume WHERE email = ?", (from_email,))
            step = cursor.fetchone()[0]
        cursor.close()
        conn.close()

        print(f"On {from_email} current step is =>{step}")
        for message in email_messages:
            msg = service.users().messages().get(
                userId='me', id=message['id']).execute()

            if 'UNREAD' in msg['labelIds']:
                filename = 'log.txt'

                # Open the file in append mode
                with open(filename, 'a') as file:
                    # Append the new content to the file
                    file.write(
                        f"From candidate => {clean_msg(msg['snippet'])}\n")

                if step == 1:
                    res = init_answer(from_email, clean_msg(msg['snippet']))
                    print(res)
                    if (res == 'Yes'):
                        set_step(from_email, 2)
                        res = JD_recruiter(from_email, False)
                        print(res)
                        send_email(from_email, res)
                    elif (res == 'No'):
                        res = reason_message()
                        print(res)
                        send_email(from_email, res)
                        set_step(from_email, 0)
                    else:
                        res = init_message(from_email, True)
                        print(res)
                        send_email(from_email, res)
                elif step == 0:
                    res = (clean_msg(msg['snippet']))
                    print(res)
                    send_email(from_email, end_message())
                    set_step(from_email, 10)
                elif step == 2:
                    res = JD_recruiter_answer(
                        from_email, clean_msg(msg['snippet']))
                    print(res)
                    if (res == 'Go'):
                        set_step(from_email, 3)
                        states[from_email] = 'Go'
                        res = show_JD(from_email)
                        print(res)
                        send_email(from_email, res)
                    elif (res == 'Interview'):
                        set_step(from_email, 3)
                        states[from_email] = "Interview"
                        timeslot[from_email] = calendar_show()
                        print(timeslot[from_email])
                        send_email(from_email, timeslot[from_email])
                    else:
                        res = JD_recruiter(from_email, True)
                        print(res)
                        send_email(from_email, res)
                elif step == 3:
                    if states[from_email] == "Go":
                        res = show_JD_answer(
                            from_email, clean_msg(msg['snippet']))
                        print(res)
                        if res == "Yes":
                            set_step(from_email, 4)
                            call_res = open_job(from_email)
                            print(call_res)
                            send_email(from_email, call_res)
                        elif res == "No":
                            call_res = end_message()
                            print(call_res)
                            send_email(from_email, call_res)
                        else:
                            call_res = show_JD(from_email)
                            print(call_res)
                            send_email(from_email, call_res)
                    else:
                        res = calendar_book(timeslot[from_email], clean_msg(
                            msg['snippet']), from_email)
                        print(res)
                        if res == "None":
                           timeslot[from_email] = calendar_show()
                           print(timeslot[from_email])
                           send_email(from_email, timeslot[from_email])
                        else:
                            call_res = reserve_message(from_email, res)
                            set_step(from_email, 10)
                            print(call_res)
                            send_email(from_email, call_res)
                            
                        
                elif step == 4:
                    res = open_job_answer(
                        from_email, clean_msg(msg['snippet']))
                    print(res)
                    if res == "Yes":
                        set_step(from_email, 5)
                        call_res = commute_job(from_email)
                        print(call_res)
                        send_email(from_email, call_res)
                    else:
                        call_res = end_message()
                        print(call_res)
                        send_email(from_email, call_res)
                        set_step(from_email, 10)
                elif step == 5:
                    res = commute_job_answer(
                        from_email, clean_msg(msg['snippet']))
                    print(res)
                    if res == "Yes":
                        set_step(from_email, 6)
                        timeslot[from_email] = calendar_show()
                        print(timeslot[from_email])
                        send_email(from_email, timeslot[from_email])
                    else:
                        call_res = commute_job(from_email)
                        print(call_res)
                        send_email(from_email, call_res)
                elif step == 6:
                    res = calendar_book(timeslot[from_email], clean_msg(
                        msg['snippet']), from_email)
                    print(res)
                    if res != "None":
                        call_res = reserve_message(
                            from_email, clean_msg(msg['snippet']))
                        print(call_res)
                        send_email(from_email, call_res)
                        set_step(from_email, 10)
                    else:
                        timeslot[from_email] = calendar_show()
                        print(timeslot[from_email])
                        send_email(from_email, timeslot[from_email])

                email_mark_as_read(message['id'])

# send complex email
def read_email():
    global states, timeslot
    email_items = []
    # get the emails from SQLite DB
    with sqlite3.connect('mydb.sqlite') as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT email FROM resume")
        email_items = [row[0] for row in cursor.fetchall()]
    cursor.close()
    conn.close()

    for from_email in email_items:

        results = service.users().messages().list(
            userId='me', q=f"from:{from_email}").execute()
        email_messages = results.get('messages', [])

        with sqlite3.connect('mydb.sqlite') as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT step FROM resume WHERE email = ?", (from_email,))
            step = cursor.fetchone()[0]
        cursor.close()
        conn.close()

        print(f"On {from_email} current step is =>{step}")
        for message in email_messages:
            msg = service.users().messages().get(
                userId='me', id=message['id']).execute()

            if 'UNREAD' in msg['labelIds']:
                filename = 'log.txt'

                # Open the file in append mode
                with open(filename, 'a') as file:
                    # Append the new content to the file
                    file.write(f"From candidate => {clean_msg(msg['snippet'])}\n")

                if step == 1:
                    res = init_answer(from_email, clean_msg(msg['snippet']))
                    print(res)
                    if (res == 'Yes'):
                        set_step(from_email, 2)
                        res = JD_recruiter(from_email, False)
                        print(res)
                        send_email(from_email, res)
                    elif (res == 'No'):
                        res = reason_message()
                        print(res)
                        send_email(from_email, res)
                        set_step(from_email, 0)
                    else:
                        res = init_message(from_email, True)
                        print(res)
                        send_email(from_email, res)
                elif step == 0:
                    res = (clean_msg(msg['snippet']))
                    print(res)
                    send_email(from_email, end_message())
                    set_step(from_email, 10)
                elif step == 2:
                    res = JD_recruiter_answer(from_email, clean_msg(msg['snippet']))
                    print(res)
                    if (res == 'Go'):
                        set_step(from_email, 3)
                        states[from_email] = 'Go'
                        res = show_JD(from_email)
                        print(res)
                        send_email(from_email, res)
                    elif (res == 'Interview'):
                        set_step(from_email, 3)
                        states[from_email] = "Interview"
                        timeslot[from_email] = calendar_show()
                        print(timeslot[from_email])
                        send_email(from_email, timeslot[from_email])
                    else:
                        res = JD_recruiter(from_email, True)
                        print(res)
                        send_email(from_email, res)
                elif step == 3:
                    if states[from_email] == "Go":
                        res = show_JD_answer(
                            from_email, clean_msg(msg['snippet']))
                        print(res)
                        if res == "Yes":
                            set_step(from_email, 4)
                            call_res = confirm_screening(
                                from_email)
                            print(call_res)
                            send_email(from_email, call_res)
                        elif res == "No":
                            call_res = end_message()
                            print(call_res)
                            send_email(from_email, call_res)
                        else:
                            call_res = show_JD(from_email)
                            print(call_res)
                            send_email(from_email, call_res)
                    else:
                        res = calendar_book(timeslot[from_email], clean_msg(msg['snippet']), from_email)
                        print(res)
                        if res == "None":
                            timeslot[from_email] = calendar_show()
                            print(timeslot[from_email])
                            send_email(from_email, timeslot[from_email])
                        else:
                            call_res = reserve_message(from_email, res)
                            set_step(from_email, 10)
                            print(call_res)
                            send_email(from_email, call_res)
                elif step == 4:
                    res = confirm_screening_answer(
                        from_email, clean_msg(msg['snippet']))
                    print(res)
                    if res == "Yes":
                        set_step(from_email, 5)
                        call_res = screening_question(from_email)
                        print(call_res)
                        send_email(from_email, call_res)
                    else:
                        call_res = other_skill_end()
                        print(call_res)
                        send_email(from_email, call_res)
                        set_step(from_email, 10)
                elif step == 5:
                    res = screening_question_answer(from_email, clean_msg(msg['snippet']))
                    print(res)
                    if res == "Yes":
                        set_step(from_email, 6)
                        call_res = question_motivate(from_email)
                        print(call_res)
                        send_email(from_email, call_res)
                    elif res == "No":
                        call_res = screening_question(from_email)
                        print(call_res)
                        send_email(from_email, call_res)
                elif step == 6:
                    print(clean_msg(msg['snippet']))
                    set_step(from_email, 7)
                    res = question_salary(from_email)
                    print(res)
                    send_email(from_email, res)
                elif step == 7:
                    print(clean_msg(msg['snippet']))
                    set_step(from_email, 8)
                    res = more_question(from_email)
                    print(res)
                    send_email(from_email, res)
                elif step == 8:
                    print(clean_msg(msg['snippet']))
                    set_step(from_email, 9)
                    timeslot[from_email] = calendar_show()
                    print(timeslot[from_email])
                    send_email(from_email, timeslot[from_email])
                elif step == 9:
                    res = calendar_book(timeslot[from_email], clean_msg(msg['snippet']), from_email)
                    print(res)
                    if res != "None":
                        call_res = reserve_message(
                            from_email, clean_msg(msg['snippet']))
                        print(call_res)
                        send_email(from_email, call_res)
                        set_step(from_email, 10)
                    else:
                        timeslot[from_email] = calendar_show()
                        print(timeslot[from_email])
                        send_email(from_email, timeslot[from_email])

                email_mark_as_read(message['id'])


@app.route('/message', methods=['POST'])
def message():
    incoming_msg = request.values.get('Body', '').lower()

    from_address = request.values.get('From')


@app.route('/resume_upload', methods=['GET', 'POST'])
def resume_upload():
    if request.method == 'POST':

        resume = request.files['resume']
        filename = resume.filename

        filename = str(uuid.uuid4()) + f".{filename.split('.')[-1]}"
        resume.save(os.path.join(f"./uploads/{filename}"))

        info = extract_info(filename)

        email_init_message(info['email'], info['name'])

        file = open("current_candidate.txt", "w")
        file.write(info['email'])
        file.close()

        return jsonify(info)


def clean_DB():
    # Connect to the database
    conn = sqlite3.connect('mydb.sqlite')
    cursor = conn.cursor()
    cursor.execute('DELETE FROM resume;')
    conn.commit()
    # Close the connection
    cursor.close()
    conn.close()

def insert_resume_db(resume_data):
    conn = sqlite3.connect('mydb.sqlite')
    cur = conn.cursor()

    first_name = resume_data['first_name']
    last_name = resume_data['last_name']
    email = resume_data['email']
    mobile = resume_data['mobile']
    filename = resume_data['filename']
    step = 1
    cur.execute('SELECT * FROM resume WHERE filename = ?', (filename,))

    row = cur.fetchone()
    if not row:
        try:
            cur.execute('INSERT INTO resume (first_name, last_name, email, mobile, filename, step) VALUES (?, ?, ?, ?, ?,?)',
                        (first_name, last_name, email, mobile, filename, step))
            conn.commit()
        except sqlite3.Error as e:
            print(f"Database error: {e}")
        except Exception as e:
            print(f"An error occurred: {e}")
        finally:
            cur.close()
            conn.close()


@app.route('/log', methods=['GET', 'POST'])
def log():
    filename = 'log.txt'
    message = None
    with open(filename, 'r') as file:
        message = file.read()

    status_code = 200
    response = make_response(message, status_code)
    response.headers['Content-Type'] = 'text/plain'
    return response

@app.route('/screen_start', methods=['POST'])
def screen_start():
    if request.method == 'POST':
        clean_DB()
        res = request.json
        print(res)
        job_id = res['job_description_title']
        level = res['level']
        with open(f'../job_description/{job_id}.json') as file:
            job_description = json.load(file)['job_description']
        for candidate in res['candidates']:
            email = candidate['email']
            number = candidate['phone']
            first_name = candidate['first_name']
            last_name = candidate['last_name']
            name = f"{first_name} {last_name}"

            # Define the data to be written
            data = {
                "job_description": job_description,
                "job_title": job_id,
                "candidate_name": name
            }

            # Open the file for writing
            with open('info.json', 'w') as file:
                # Write the data to the file in JSON format
                json.dump(data, file)

            # if not '+' in number:
            #     number = '+' + number

            # phone_message(f"Hi, {name}\n{screen_question}", number, 'sms')
            # phone_message(f"Hi, {name}\n{screen_question}", number, 'wa')
            # end

            resume_data = {}
            resume_data['first_name'] = first_name
            resume_data['last_name'] = last_name
            resume_data['email'] = email
            resume_data['mobile'] = number

            resume_data['filename'] = f"{first_name} {last_name}_{number}"
            func_res = init_message(email, False)
            send_email(email, func_res)
            # set step in DB
            insert_resume_db(resume_data)
        while (True):
            if level == "complex":
                read_email()
            elif level == "medium":
                read_md_email()
            else:
                read_simple_email()
            time.sleep(20)

    
    return "Good"


if __name__ == '__main__':
    app.run(debug=True, port=5000, host=('0.0.0.0'))
