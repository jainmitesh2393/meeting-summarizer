import speech_recognition as sr
import pyttsx3
from datetime import datetime

f=open("TASKS.docx", "a")
f.write("HERE ARE YOUR TASKS FOR THE DAY")
# Initialize the recognizer and text-to-speech engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()

# Dictionary to store tasks with their details
tasks = {}
task_name={}

# Function to speak text
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Function to listen to user's voice and return text
def listen():
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        print("Listening...")
        audio = recognizer.listen(source)
        try:
            text = recognizer.recognize_google(audio)
            print(f"You said: {text}")
            return text.upper()
        except sr.UnknownValueError:
            print("Sorry, I didn't catch that.")
            speak("Sorry, I didn't catch that.")
            return ""
        except sr.RequestError:
            print("Network error.")
            speak("Network error.")
            return ""

# Function to get a date and time input from the user
def get_datetime():
    """speak("Please say the due date in the format: day, month, and year.")
    date_str = listen()
    speak("Please say the due time in the format: hour and minutes.")
    time_str = listen()

    try:
        due_date = datetime.strptime(date_str + " " + time_str, "%d %B %Y %H %M")
        return due_date
    except ValueError:
        speak("Invalid date or time format. Please try again.")
        return get_datetime()"""
    speak("Please say the due date in the format: day, month, and year.")
    print("Enter the due date (in DD M YYYY format):")
    date_str = listen()  # Listen for '26 9 2024' (space-separated format)
    speak("Please say the due time in the format: hour and minutes.")
    print("Enter the due time (in HHMM format):")
    time_str= listen()  # Listen for '1100'

    # Handle date format "DD M YYYY"
    try:
        due_date_parsed = datetime.strptime(date_str, "%d %m %Y")
    except ValueError:
        #print("Invalid date format. Please use 'DD M YYYY'.")
        return
    
    # Handle time format "HHMM"
    try:
        due_time_parsed = datetime.strptime(time_str, "%H%M")
    except ValueError:
        print("Invalid time format. Please use 'HHMM'.")
        return

    print(f"Task '{task_name}' created with due date {due_date_parsed.strftime('%d/%m/%Y')} and time {due_time_parsed.strftime('%H:%M')}.")
    f.write("\nTask", task_name,"\n\tDue Date: ", due_date_parsed, "\tTime: ",due_time_parsed)
    #f.write("\nTask '{task_name}' created with due date {due_date_parsed.strftime('%d/%m/%Y')} and time {due_time_parsed.strftime('%H:%M')}.")
    return date_str


# Function to create a new task
def create_task():
    speak("What is the task?")
    print("What is the task?")
    task_name = listen()
    if task_name:
        due_date = get_datetime()
        speak("Do you want to set a reminder? Say yes or no.")
        
        print("Do you want to set a reminder? Say yes or no.")
        reminder = listen()

        task = {
            "name": task_name,
            "due_date": due_date,
            "reminder": True if "yes" in reminder else False,
            "subtasks": []
        }
        tasks[task_name] = task
        speak(f"Task '{task_name}' added successfully.")
        print(f"Task '{task_name}' added successfully.")

# Function to edit an existing task
def edit_task():
    if not tasks:
        speak("No tasks available to edit.")
        print("No tasks available to edit.")
        return

    speak("Which task do you want to edit?")
    print("Which task do you want to edit?")

    task_name = listen()

    if task_name in tasks:
        speak(f"Editing task: {task_name}. What do you want to edit? Say due date, due time, or reminder.")
        print(f"Editing task: {task_name}. What do you want to edit? Say due date, due time, or reminder.")
        edit_choice = listen()

        if "due date" in edit_choice or "due time" in edit_choice:
            tasks[task_name]['due_date'] = get_datetime()
            speak(f"Due date updated for task {task_name}.")
            print(f"Due date updated for task {task_name}.")
        elif "reminder" in edit_choice:
            speak("Do you want to enable or disable the reminder? Say enable or disable.")
            print("Do you want to enable or disable the reminder? Say enable or disable.")
            reminder_choice = listen()
            tasks[task_name]['reminder'] = True if "enable" in reminder_choice else False
            speak(f"Reminder updated for task {task_name}.")
            print(f"Reminder updated for task {task_name}.")
        else:
            speak("Invalid option.")
            print("Invalid option.")
    else:
        speak(f"Task '{task_name}' not found.")
        print(f"Task '{task_name}' not found.")

# Function to add a subtask
def add_subtask():
    if not tasks:
        speak("No tasks available to add a subtask.")
        print("No tasks available to add a subtask.")

        return

    speak("Which task do you want to add a subtask to?")
    print("Which task do you want to add a subtask to?")

    task_name = listen()

    if task_name in tasks:
        speak(f"What is the subtask for task {task_name}?")
        print(f"What is the subtask for task {task_name}?")
        subtask_name = listen()
        if subtask_name:
            tasks[task_name]['subtasks'].append(subtask_name)
            speak(f"Subtask '{subtask_name}' added to task '{task_name}'.")
            print(f"Subtask '{subtask_name}' added to task '{task_name}'.")
    else:
        speak(f"Task '{task_name}' not found.")
        print(f"Task '{task_name}' not found.")

# Function to display menu and get user choice
def get_menu_choice():
    speak("Please choose an option: create a task, edit a task, add a subtask, display tasks or say exit to quit.")
    print("1. Create a Task\n2. Edit a Task\n3. Add a subtask\n4. Display Tasks\nExit")
    return listen()

# Main loop to run the to-do list
def todo_list():
    print("Welcome to your voice-controlled to-do list.\n Please speak after the 'Listening...' prompt appears")
    speak("Welcome to+. your voice-controlled to-do list.")
    while True:
        choice = get_menu_choice()

        if "create" in choice:
            create_task()
        elif "edit" in choice:
            edit_task()
        elif "subtask" in choice or "add subtask" in choice:
            add_subtask()
        elif "exit" in choice:
            f.close()
            speak("Goodbye!")
            break
        elif "display" or "print" or "show" in choice:
            speak("Here are your tasks.")
            print(tasks)

        else:
            speak("Invalid choice, please try again.")
            print("Invalid choice, please try again.")
        f.close()


# Run the to-do list program
if __name__ == "__main__":
    todo_list()
