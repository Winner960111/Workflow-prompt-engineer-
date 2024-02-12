import sqlite3
from O365 import Account, MSGraphProtocol
from datetime import datetime, timedelta
import datetime as dt
from openai import OpenAI
import os
import json
from dotenv import load_dotenv
load_dotenv()

openai_api_key = os.environ.get("OPENAI_API_KEY")
openai_client = OpenAI(api_key=openai_api_key)
chat_message = {}
# generate message using chatGTP

with open("info.json", 'r') as file:
    data_des = json.load(file)
job_title = data_des['job_title']
job_description = data_des['job_description']

calendar_tools = [
    {
        "type": "function",
        "function": {
            "name": "get_time_slot",
            "description": "Get the selected time slot",
            "parameters": {
                "type": "object",
                "properties": {
                    "start_datetime": {
                        "type": "string",
                        "description": "Starting date and time of selected time slot.",
                    },
                    "end_datetime": {
                        "type": "string",
                        "description": "Ending date and time of selected time slot.",
                    },
                },
                "required": ["start_datetime", "end_datetime"],
            },
        }
    }
]


def chatbot(system, user, email):
    messages = []
    global chat_message

    # Add each new message to the list
    messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": user})
    # Request gpt-4 for chat completion
    response = openai_client.chat.completions.create(
        model="gpt-4",
        messages=messages
    )

    # Print the response and add it to the messages list
    temp = response.choices[0].message.content
    while '[' in temp and ']' in temp:
        start = temp.find('[')
        end = temp.find(']', start)
        if start != -1 and end != -1:
            temp = temp[:start] + temp[end+1:]
        else:
            break
    chat_message[email] = temp
    return chat_message[email]

# chapter 1: init_message


def init_message(email, other):
    with open("info.json", 'r') as file:
        data_des = json.load(file)
        candidate_name = data_des['candidate_name']
        
    system_content = f"As a professional AI recruitment assistant from Portal, generate a professional and polite recruitment outreach message for an job description and job title to a candidate, highlighting that candidate's profile is a strong match for the role at the location of the office, and inquiring about her interest in learning more about the opportunity. You should remeber that job description is {job_description}, job title is {job_title} and and candidate's name is {candidate_name}. You should make simply sentences inside 30 words. You don't need to attach subject, your name, or your position on message. If {other} equeal to 'True', you don't need to share greeting words and candidate's name."

    user_content = "Please send the message to candidate"
    res = chatbot(system_content, user_content, email)
    return res

# analyze init_answer


def init_answer(email, input):
    system_content = f"As a information analyzer, you should answer only a word. Answers as 'Yes' if {input} is positive for {chat_message[email]} and answers as 'No' if {input} is  for {chat_message[email]}. Also answers 'other' if {input} is neither positive nor negative for {chat_message[email]}"
    user_content = f"please analyze {input}"
    res = chatbot(system_content, user_content, email)
    return res

# step JD or connect with recruiter


def JD_recruiter(email, other):
    system_content = f"As a professional AI recruitment assistant from Portal, generates messages asking whether candidates want to see a job description or be interviewed directly by a hiring manager. You must create a simple sentence of 30 words or less. Greetings should not be shared. You should not always attach a subject, your name, or your position."
    user_content = "Please send the message to candidate"
    res = chatbot(system_content, user_content, email)
    return res

# analyze JD_recruiter


def JD_recruiter_answer(email, input):
    system_content = f"As an information analyzer, you only need to answer one word. If the input content is a request for a job description of {job_description}, respond with 'Go'. And if the input content is a request for a direct interview with the hiring manager, respond with 'Interview'. If the information entered is not a job description or interview with a recruiter, respond with 'Other'. The information entered here is {input}."
    user_content = input
    res = chatbot(system_content, user_content, email)
    return res

# show candidate the JD


def show_JD(email):
    system_content = f"As a professional AI recruitment assistant from Portal, you should provide job description as following format:'Job description:\n {job_description}', and ask that Can candidate share how he believe his skill set aligns with the requirements and demands of our company. Job description is {job_description}. you should show the job description without adding or removing any contents. You should not always attach topics, your name, or your position."
    user_content = "Please send the message to candidate"
    res = chatbot(system_content, user_content, email)
    return res
# analyze show_JD answer


def show_JD_answer(email, input):
    system_content = f"As a information analyzer, you should answer only a word. If {input} is positive, respond with 'Yes' and is negative, respond with 'No'."
    user_content = f"please analyze {input}"
    res = chatbot(system_content, user_content, email)
    return res


# connect candidates with recruiter manager
def connect_interview(email):
    system_content = f"As a professional AI recruitment assistant from Portal, you should ask if when available the time slot of the candidate to him inside today or tomorrow. You should make simply sentences inside 30 words. You should not always share topics, your name, or your position."
    user_content = "Please send the message to candidate"
    res = chatbot(system_content, user_content, email)
    return res

# analyze connect_interview answer


def connect_interview_answer(email, input):
    system_content = f"As a information analyzer, you should answer only a word. If {input} is content meaning the time, respond with 'Time'. If {input} isn't content meaning the time, respond with 'Other'."
    user_content = f"please analyze {input}"
    res = chatbot(system_content, user_content, email)
    return res

# suggest the question that confirm if the candidate fit for this role


def confirm_screening(email):
    chat_message[email] = "Excellent! I would like to know some experience & Skills required for these Role."
    return chat_message[email]


def confirm_screening_answer(email, input):
    system_content = f"As a information analyzer, you should answer only a word. If {input} is related to {job_description}, respond with 'Yes' and is negative, respond with 'No'."
    user_content = f"please analyze {input}"
    res = chatbot(system_content, user_content, email)
    return res

#ask open for this job
def open_job(email):
    chat_message[email] = "good! And I would like to know if you have a vacancy for this job."
    return chat_message[email]

def open_job_answer(email, input):
    system_content = f"As a information analyzer, you should answer only a word. If {input} is positive about {chat_message[email]}, respond with 'Yes' and is negative, respond with 'No'."
    user_content = f"please analyze {input}"
    res = chatbot(system_content, user_content, email)
    return res
# ask commute for this job

def commute_job(email):
    system_content = f"As a professional AI recruitment assistant from Portal, you should ask if the candidate can commute and how far does the candidate live from the location of the company. Consider the location of the company from {job_description}. You should make simply sentences inside 30 words. You should not always share topics, your name, or your position."
    user_content = "Please send the message to candidate"
    res = chatbot(system_content, user_content, email)
    return res

def commute_job_answer(email, input):
    system_content = f"As a information analyzer, you should answer only a word. If {input} is related to {chat_message[email]} wether it's posive state or is negative one, respond with 'Yes' and is not related, respond with 'No'."
    user_content = f"please analyze {input}"
    res = chatbot(system_content, user_content, email)
    return res

# screening question

def screening_question(email):
    system_content = f"As a professional AI recruitment assistant from Portal, You should make 3 screening question based on {job_description}. The questions of experiences and cirtification for {job_description} must be included in screening question. Each screening question should be inside 20 words."
    user_content = "Please send the message to candidate"
    res = chatbot(system_content, user_content, email)
    return res


def screening_question_answer(email, input):
    system_content = f"As a information analyzer, you should answer only a word. Answer 'Yes' if {input} is related to {chat_message[email]} and answer 'No' if {input} is related to {chat_message[email]}."
    user_content = f"please analyze {input}"
    res = chatbot(system_content, user_content, email)
    return res

# suggest motivate question


def question_motivate(email):
    system_content = f"As a professional AI recruitment assistant from Portal, You've recently met with a candidate who has a notable record of achievements in their current role but has decided to apply for a new position at your company. As part of the interview process, you want to understand what is driving their desire to change jobs and what they are looking for in new opportunities. Generate a message that encourages the candidate to share their motivations and reasons behind this transition. Greetings should not be shared. You should make simply sentences inside 30 words."
    user_content = "Please send the message to candidate"
    res = chatbot(system_content, user_content, email)
    return res

# suggest salary question


def question_salary(email):
    system_content = f"As a professional AI recruitment assistant from Portal, During a job interview, after discussing the candidate's experience and skills, it is important to address compensation expectations. Compose a professional question that transitions from the prior topic of conversation and inquires about the candidate's salary expectations for the position they are interviewing for. You should make simply sentences inside 30 words."
    user_content = "Please send the message to candidate"
    res = chatbot(system_content, user_content, email)
    return res

# suggest more question


def more_question(email):
    chat_message[email] = " Perfect, I appreciate your flexibility. Lastly, as part of our hiring process, there will be a leadership review call. Are you comfortable with that, and do you have any questions or concerns about the upcoming call?"
    return chat_message[email]

# reserve interview


def reserve_message(email, input):
    system_content = f"As a professional AI recruitment assistant from Portal, you should schedule interview between candidate and recruitment manager based on the time of {input}. you should make simply sentences within 20 words. You don't need to share the subject and your name."
    user_content = "Please send the message to candidate"
    res = chatbot(system_content, user_content, email)
    return res


def reason_message():
    return ("Could you please list out the reason, why you're not looking out for a change.")
# end


def end_message():
    return ("Thanks for letting us know, Will keep in touch for future. Have a Great Day.")


def other_skill_end():
    return ("Sorry but, your skills aren't fit for our role. Thanks for your interest.")

# reserve timeslot


def calendar_show():

    # Establish a connection to the database
    conn = sqlite3.connect('mydb.sqlite')
    cursor = conn.cursor()

    # Execute a query to retrieve specific columns
    cursor.execute('SELECT start_time, end_time FROM resume')

    # Fetch all rows from the query result
    rows = cursor.fetchall()

    # Convert the rows into a list of lists, where each inner list represents a row
    busy_events = [dict(start_time=row[0], end_time=row[1]) for row in rows]
    # Close the cursor and connection
    cursor.close()
    conn.close()
    # print(f"busy event is recieved as following:{busy_events}")

    start_time = datetime.now().replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
    end_time = start_time + timedelta(days=1)
    available_slots = []
    current_time = start_time

    while len(available_slots) < 3 and current_time < end_time:
        format = "%Y-%m-%dT%H:%M:%SZ"
        slot_end_time = current_time + timedelta(minutes=30)
        for event in busy_events:
            # Use square brackets for dictionary key access
            start = event['start_time']
            # Use square brackets for dictionary key access
            end = event['end_time']

            if start == None and end == None:
                slot_busy = False
            elif datetime.strptime(start, format) >= current_time and datetime.strptime(end, format) <= slot_end_time:
                slot_busy = True
            else:
                slot_busy = False
            if slot_busy:
                break
        if not slot_busy:
            available_slots.append((current_time, slot_end_time))
        current_time += timedelta(minutes=30)
    content = ""
    for slot in available_slots:
        content += f"{slot[0]} - {slot[1]}\n"
    messages = []
    messages.append(
        {"role": "system", "content": "As a professional AI assistant, you should speak polite and kind. you should ask shortly inside 20 words include timeslot. You should display the list of timeslot. You must consider current date."})
    messages.append(
        {"role": "user", "content": f"please make the message that ask to candidate available interview time based on {content}."})

    response = openai_client.chat.completions.create(
        model="gpt-4-1106-preview",
        messages=messages
    )
    res = f"Good! And then, I would like to schedule an interview with the hiring manager for this project.\n{response.choices[0].message.content}"
    return res


def calendar_book(bot_msg, candidate_msg, email):
    messages = []
    messages.append(
        {"role": "system", "content": f"Please answer which time slot is selected by user with current date. You should only answer as time. {bot_msg} include the current date."})
    messages.append(
        {"role": "user", "content": f"{bot_msg}\nUser's response: {candidate_msg}"})

    print(f"this is bot's message {bot_msg}")
    response = openai_client.chat.completions.create(
        model="gpt-4-1106-preview",
        messages=messages,
    )

    set_time = response.choices[0].message.content
    messages = []
    messages.append(
        {"role": "system", "content": "Please answer which time slot is selected by user. If user didn't select time slot, you have to answer as None"})
    messages.append({"role": "user", "content": set_time})

    response = openai_client.chat.completions.create(
        model="gpt-4-1106-preview",
        messages=messages,
        tools=calendar_tools,
    )

    try:
        function_time = response.choices[0].message.tool_calls[0].function.arguments
        start_time = function_time.split("\"")[3]
        end_time = function_time.split("\"")[7]
        if start_time[-1] != 'Z':
            start_time+='Z'
        if end_time[-1] != 'Z':
            end_time += 'Z'
        print(f"save time format => {start_time}")
        # Connect to the database
        conn = sqlite3.connect('mydb.sqlite')
        cur = conn.cursor()

        cur.execute('SELECT * FROM resume WHERE email = ?', (email, ))
        row = cur.fetchone()
        
        if row != None:
            cur.execute(
                "UPDATE resume SET start_time = ?, end_time = ? WHERE email = ?", (start_time, end_time, email))
            conn.commit()
            cur.close()
            conn.close()
        else:
            print("Error")
        interviewTime = f"{start_time} + {end_time}"
        print(f"interviewTime is => {interviewTime}")
        return interviewTime
    except:
        return "None"
