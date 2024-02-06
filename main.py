from O365 import Account, MSGraphProtocol
import datetime as dt
from datetime import datetime, timedelta
from openai import OpenAI

openai_client = OpenAI(api_key="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")

CLIENT_ID = 'cd3d12ee-6bb3-437c-b613-37caa4ee398e'
SECRET_ID = 'nog8Q~JIEzUIfHO2WRILResZOzNUXnNjoHF6Lcli'

credentials = (CLIENT_ID, SECRET_ID)

protocol = MSGraphProtocol(default_resource='') 
#protocol = MSGraphProtocol(defualt_resource='<sharedcalendar@domain.com>') 
scopes = ['Calendars.Read.Shared']
calendar_scopes = protocol.get_scopes_for('calendar_all')
scopes.extend(calendar_scopes)

account = Account(credentials, protocol=protocol)

if account.authenticate(scopes=scopes):
   print('Microsoft Calendar Authenticated!')

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

def calendar_show():
    schedule = account.schedule()
    calendar = schedule.get_default_calendar()

    start_time = datetime.now().replace(hour=9, minute=0, second=0, microsecond=0)
    end_time = start_time + timedelta(days=1)

    busy_events = calendar.get_events(include_recurring=False)

    available_slots = []
    current_time = start_time
    while len(available_slots) < 3 and current_time < end_time:
        slot_end_time = current_time + timedelta(minutes=30)
        slot_busy = any(event for event in busy_events if event.start <= current_time and event.end >= slot_end_time)
        if not slot_busy:
            available_slots.append((current_time, slot_end_time))
        current_time += timedelta(minutes=30)

    content = "Please select one available meeting time slots among following three options.\n\n"
    
    for slot in available_slots:
        content += f"{slot[0]} - {slot[1]}\n"

def calendar_book(bot_msg, candidate_msg):
    schedule = account.schedule()
    calendar = schedule.get_default_calendar()
    messages = []
    messages.append({"role": "system", "content": "Please answer which time slot is selected by user."})
    messages.append({"role": "user", "content": f"{bot_msg}\n{candidate_msg}"})

    response = openai_client.chat.completions.create(
        model="gpt-4-1106-preview",
        messages=messages,
    )

    set_time = response.choices[0].message.content

    messages = []
    messages.append({"role": "system", "content": "Please answer which time slot is selected by user. If user didn't select time slot, you have to answer as None"})
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

        format = "%Y-%m-%dT%H:%M:%S"

        start_dt = datetime.strptime(start_time, format)
        end_dt = datetime.strptime(end_time, format)

        new_event = calendar.new_event()  # creates a new unsaved event
        new_event.subject = 'Interview with Candidate'
        new_event.location = 'UTC time Zone'

        # naive datetimes will automatically be converted to timezone aware datetime
        #  objects using the local timezone detected or the protocol provided timezone

        new_event.start = start_dt
        new_event.end = end_dt
        # so new_event.start becomes: datetime.datetime(2018, 9, 5, 19, 45, tzinfo=<DstTzInfo 'Europe/Paris' CEST+2:00:00 DST>)

        # new_event.recurrence.set_daily(1, end=datetime(2024, 2, 4))
        # new_event.remind_before_minutes = 45

        new_event.save()

    except:
        calendar_show()
    
