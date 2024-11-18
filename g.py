import pandas as pd
from openpyxl import Workbook
import streamlit as st
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import speech_recognition as sr
import pyttsx3
import time
from datetime import datetime
from groq import Groq

# Initialize the recognizer and text-to-speech engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()

# Function to speak text
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Function to listen to user's voice and return text
def listen():
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        st.info("Listening for the meeting summary... Please speak clearly.")
        audio = recognizer.listen(source)
        try:
            text = recognizer.recognize_google(audio)
            st.success(f"You said: {text}")
            return text.lower()
        except sr.UnknownValueError:
            st.warning("Sorry, I didn't catch that.")
            speak("Sorry, I didn't catch that.")
            return ""
        except sr.RequestError:
            st.error("Network error. Please check your connection.")
            speak("Network error.")
            return ""

# Function to display tasks or MoM content in sticky notes
def display_sticky_notes(content, is_mom=False):
    items = content.split('\n\n')  # Split items by double new lines
    for item in items:
        if item.strip():  # Only display non-empty items
            st.markdown(f"""
                <div style="
                    background-color: #F8F9FA; 
                    padding: 15px; 
                    margin-bottom: 10px; 
                    border-radius: 10px; 
                    box-shadow: 2px 2px 5px rgba(0,0,0,0.1); 
                    font-family: Arial; 
                    font-size: 16px;
                    border-left: 5px solid {'#28A745' if is_mom else '#007BFF'};
                    color: #333;
                ">
                    {item.replace('Task ID:', '<b>Task ID:</b>')
                         .replace('Assigned to:', '<b>Assigned to:</b>')
                         .replace('the due date:', '<b>Due Date:</b>')
                         .replace('the due time:', '<b>Due Time:</b>')}
                </div>
            """, unsafe_allow_html=True)

# Function to initialize Groq API client
def initialize_groq_client():
    api_key = "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"  # Replace with your actual API key
    try:
        return Groq(api_key=api_key)
    except Exception as e:
        st.error(f"Error initializing Groq client: {e}")
        return None

# Function to generate a to-do list using the Groq API
def generate_todo_list(client, meeting_summary):
    system_prompt = """
    You are a smart assistant that generates to-do lists from meeting summaries. Your task is to read a paragraph that summarizes a meeting and extract specific action items into a concise to-do list. Each task must be formatted as follows:

    Task ID: <unique ID in ascending order>
    <task description>
    Assigned to: <person responsible for the task>
    the due date (in DD M YYYY format): <due date or "N/A" if not mentioned>
    the due time (in HHMM format): <due time or "N/A" if not mentioned>
    """

    conversation = f"Meeting Summary: {meeting_summary}\nAssistant: Please extract the tasks, create a to-do list with unique task IDs, assigned persons, due dates, and times."

    try:
        chat_completion = client.chat.completions.create(
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": conversation}
            ],
            model="llama3-70b-8192",
            temperature=0.5
        )
        response = chat_completion.choices[0].message.content
        return response.encode('utf-8').decode('utf-8')
    except Exception as e:
        st.error(f"Error generating to-do list: {e}")
        return "An error occurred while generating the to-do list."

# Function to generate MoM format from a to-do list
def generate_mom_format(todo_list):
    tasks = todo_list.split('\n\n')
    mom_output = "MINUTES OF MEETING\n\nDate: [Insert Date Here]\nVenue: [Insert Venue Here]\nParticipants: [Insert Participants Here]\n\n"
    mom_output += "Sr. No.\nPoints Raised / Suggested\nTimelines\nResponsibility\nRemarks\n\n"

    for idx, task in enumerate(tasks, start=1):
        if task.strip():
            lines = task.split('\n')
            description = lines[1] if len(lines) > 1 else "N/A"
            assigned_to = lines[2].replace("Assigned to:", "").strip() if len(lines) > 2 else "N/A"
            due_date = lines[3].replace("the due date:", "").strip() if len(lines) > 3 else "N/A"
            remarks = "Remarks: [Insert Remarks Here]"

            mom_output += f"{idx}\n{description}\n{due_date}\n{assigned_to}\n{remarks}\n\n"

    return mom_output

# Function to parse MoM format into a structured dictionary
def parse_mom_to_dict(mom_text):
    tasks = mom_text.strip().split("Sr. No.")
    parsed_data = []

    for task in tasks[1:]:  # Skip the header
        lines = task.strip().splitlines()
        if len(lines) >= 4:
            parsed_data.append({
                "Sr. No.": lines[0].strip(),
                "Description": lines[1].strip(),
                "Timeline": lines[2].replace("Timeline:", "").strip(),
                "Responsibility": lines[3].replace("Responsibility:", "").strip(),
                "Remarks": lines[4].replace("Remarks:", "").strip() if len(lines) > 4 else ""
            })

    return parsed_data

# Function to save parsed MoM data to an Excel file
def mom_to_excel(mom_text, output_filename="MoM_Output.xlsx"):
    parsed_data = parse_mom_to_dict(mom_text)
    df = pd.DataFrame(parsed_data)
    df.to_excel(output_filename, index=False)
    return output_filename


# Function to send email with an attachment
def send_email_with_attachment(to_emails, subject, body, attachment_path):
    from_email = "harsh.22210267@viit.ac.in"  # Replace with your email
    password = "cgha dxgb flcq qgli"  # Use app-specific password

    for email in to_emails:
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        try:
            with open(attachment_path, "rb") as attachment:
                part = MIMEApplication(attachment.read(), Name=attachment_path)
                part['Content-Disposition'] = f'attachment; filename="{attachment_path}"'
                msg.attach(part)

            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(from_email, password)
            server.sendmail(from_email, email, msg.as_string())
            server.quit()

            st.success(f"Email successfully sent to {email}")
        except Exception as e:
            st.error(f"Error sending email: {e}")

def main():
    st.set_page_config(page_title="Meeting Manager", page_icon="📝", layout="wide")

    st.title("📝 Meeting Manager")
    st.subheader("Convert meeting summaries into actionable to-do lists and MoMs")

    if 'meeting_summary' not in st.session_state:
        st.session_state['meeting_summary'] = ""

    col1, col2 = st.columns([1, 3])

    # Input options for the meeting summary
    with col1:
        st.write("## Input Options")
        input_method = st.radio("How would you like to input your meeting summary?", ("Text", "Microphone"))
        if input_method == "Microphone":
            if st.button("🎤 Record Meeting Summary"):
                with st.spinner("Recording..."):
                    st.session_state['meeting_summary'] = listen()
        else:
            st.session_state['meeting_summary'] = st.text_area("Enter Meeting Summary:", st.session_state['meeting_summary'])
        

    # Actions and output display
    with col2:
        st.write("## Actions")

        # Generate To-Do List
        if st.button("Generate To-Do List"):
            client = initialize_groq_client()
            if client:
                todo_list = generate_todo_list(client, st.session_state['meeting_summary'])
                st.session_state['todo_list'] = todo_list
                st.success("To-Do List Generated!")
                display_sticky_notes(todo_list)

        # Generate MoM and save to Excel
        if st.button("Generate MoM"):
            if 'todo_list' in st.session_state:
                mom_content = generate_mom_format(st.session_state['todo_list'])
                st.session_state['mom_content'] = mom_content
                st.success("MoM Generated!")
                display_sticky_notes(mom_content, is_mom=True)

                # Save to Excel
                excel_file = mom_to_excel(mom_content)
                st.success(f"MoM saved to Excel: {excel_file}")

        # Email the MoM
        if st.button("📧 Email MoM"):
            default_emails = ["harsh.22210267@viit.ac.in"]  # Add recipient emails
            # Generate the MoM content from the to-do list
            client_groq_2= initialize_groq_client()
            todo_list_1= generate_todo_list(client_groq_2)
            mom_content = generate_mom_format(todo_list_1)
            excel_file = mom_to_excel(mom_content, "MoM_Output.xlsx")  # Convert MoM text to Excel
            send_email_with_attachment(default_emails, "Meeting MoM", "Please find the MoM attached.", excel_file)
if __name__ == "__main__":
    main()
