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

# Function to send the email with an attachment
def send_email_with_attachment(to_emails, subject, body, attachment_path):
    from_email = "harsh.22210267@viit.ac.in"  # Replace with your email
    password = "cgha dxgb flcq qgli"  # Replace with your app password

    for email in to_emails:
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        # Attach the Excel file
        try:
            with open(attachment_path, "rb") as attachment:
                part = MIMEApplication(attachment.read(), Name="MoM_Output.xlsx")
                part['Content-Disposition'] = 'attachment; filename="MoM_Output.xlsx"'
                msg.attach(part)
            
            st.info("Connecting to SMTP server...")
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(from_email, password)
            st.success(f"Sending email to {email}...")

            # Send email
            server.sendmail(from_email, email, msg.as_string())
            server.quit()
            st.success(f"Email successfully sent to {email}")

        except FileNotFoundError:
            st.error(f"The file {attachment_path} was not found.")
        except smtplib.SMTPException as e:
            st.error(f"SMTP error: {e}")


# Function to display tasks in light-colored sticky notes style
def display_sticky_notes(todo_list):
    tasks = todo_list.split('\n\n')  # Split tasks by double new lines
    for task in tasks:
        if task.strip():  # Only display non-empty tasks
            st.markdown(f"""
                <div style="
                    background-color: #F8F9FA; 
                    padding: 15px; 
                    margin-bottom: 10px; 
                    border-radius: 10px; 
                    box-shadow: 2px 2px 5px rgba(0,0,0,0.1); 
                    font-family: Arial; 
                    font-size: 16px;
                    border-left: 5px solid #007BFF;
                    color: #333;
                ">
                    {task.replace('Task ID:', '<b>Task ID:</b>').replace('Assigned to:', '<b>Assigned to:</b>').replace('the due date:', '<b>Due Date:</b>').replace('the due time:', '<b>Due Time:</b>')}
                </div>
            """, unsafe_allow_html=True)

# Define the generate_mom_format function to convert the to-do list into MoM format
def generate_mom_format(todo_list):
    # Split tasks by double new lines
    tasks = todo_list.split('\n\n')
    mom_output = "MINUTES OF MEETING\n\nDate: [Insert Date Here]\n\n"
    mom_output += "Venue: [Insert Venue Here]\n\n"
    mom_output += "Participants: [Insert Participants Here]\n\n"
    mom_output += "Topics Discussed:\n\n"
    
    for idx, task in enumerate(tasks, start=1):
        if task.strip():  # Only display non-empty tasks
            # Extract task details by expected format for unique ID, description, assigned person, due date, and time
            lines = task.split('\n')
            task_id = f"Sr. No. {idx}"
            description = lines[1] if len(lines) > 1 else "N/A"
            assigned_to = lines[2].replace("Assigned to:", "").strip() if len(lines) > 2 else "N/A"
            due_date = lines[3].replace("the due date:", "").strip() if len(lines) > 3 else "N/A"
            remarks = "Remarks: [Insert Remarks Here]"
            
            # Format as per MoM style
            mom_output += f"{task_id}\n"
            mom_output += f"{description}\n"
            mom_output += f"Timeline: {due_date}\n"
            mom_output += f"Responsibility: {assigned_to}\n"
            mom_output += f"{remarks}\n\n"
    
    return mom_output


# Display MoM format in Streamlit
def display_mom_format(todo_list):
    mom_content = generate_mom_format(todo_list)
    st.text_area("Minutes of Meeting Format", mom_content, height=500)

# Add a new function to parse MoM format into a DataFrame and save to Excel
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

# Function to create Excel file from MoM text
def mom_to_excel(mom_text, output_filename="MoM_Output.xlsx"):
    parsed_data = parse_mom_to_dict(mom_text)
    df = pd.DataFrame(parsed_data)
    df.to_excel(output_filename, index=False)
    return output_filename  # Return the filename for further use

# Integration with Streamlit to generate and email MoM
def generate_and_email_mom():
    meeting_summary = st.session_state.get('meeting_summary', '')
    if not meeting_summary.strip():
        st.error("Please enter or record a valid meeting summary.")
        return

    # Initialize Groq client and generate MoM format text
    client_groq = initialize_groq_client()
    todo_list = generate_todo_list(client_groq, meeting_summary)
    mom_content = generate_mom_format(todo_list)  # Convert to MoM format
    excel_file = mom_to_excel(mom_content)  # Save MoM as Excel



# Main function for the Streamlit app
def main():
    st.set_page_config(page_title="Meeting To-Do List Generator", page_icon="📝", layout="wide")

    # Page Header
    st.title("📝 Meeting To-Do List Generator")
    st.subheader("Easily convert your meeting summaries into actionable to-do lists")

    # Initialize session state for meeting summary
    if 'meeting_summary' not in st.session_state:
        st.session_state['meeting_summary'] = ""  # Initialize session state variable

    col1, col2 = st.columns([1, 3])

    with col1:
        st.write("## Input Options")
        use_microphone = st.radio(
            "How would you like to input your meeting summary?",
            ("Use Text", "Use Microphone")
        )

        if use_microphone == "Use Microphone":
            if st.button("🎤 Record Meeting Summary"):
                with st.spinner("Recording... Please speak into the microphone"):
                    time.sleep(1)
                    meeting_summary = listen()
                    if meeting_summary:
                        st.session_state['meeting_summary'] = meeting_summary  # Store in session state
                        st.write(f"Captured Meeting Summary: {meeting_summary}")
            else:
                st.write("Click 'Record' to capture a meeting summary.")
        else:
            # Ensure the text area works properly
            meeting_summary = st.text_area(
                "Enter your meeting summary:", 
                st.session_state.get('meeting_summary', ''),
                height=150
            )
            if st.button("Save Summary"):
                st.session_state['meeting_summary'] = meeting_summary  # Store in session state


    with col2:
        st.write("## To-Do List Output")
        if st.button("📝 Generate To-Do List"):
            meeting_summary = st.session_state.get('meeting_summary', '')
            if not meeting_summary.strip():
                st.error("Please enter or record a valid meeting summary.")
                return

            # Initialize Groq API client
            client_groq = initialize_groq_client()
            if client_groq is None:
                st.error("Failed to initialize the Groq client. Please check your API key.")
                st.stop()

            # Generate To-Do list from summary
            with st.spinner("Generating To-Do List..."):
                todo_list = generate_todo_list(client_groq, meeting_summary)
                if todo_list:
                    st.success("To-Do List Generated!")
                    display_sticky_notes(todo_list)  # Display each task in a sticky note style

                    # Send the to-do list via email

                else:
                    st.error("Failed to generate the to-do list.")
        
        # Replace or add this function call in the Streamlit app to display in MoM format
        if st.button("📝 Generate MOM Format"):
            meeting_summary = st.session_state.get('meeting_summary', '')
            if not meeting_summary.strip():
                st.error("Please enter or record a valid meeting summary.")
            else:
                client_groq = initialize_groq_client()
                todo_list = generate_todo_list(client_groq, meeting_summary)
                display_mom_format(todo_list)

        
        # Initialize Groq API client
        client_groq_1 = initialize_groq_client()
        if client_groq_1 is None:
            st.error("Failed to initialize the Groq client. Please check your API key.")
            st.stop()
        list = generate_todo_list(client_groq_1, meeting_summary)

        if st.button("Generate and Email MoM"):
            generate_and_email_mom()


        if st.button("📧 Send MoM via Email"):
            default_emails = ["harsh.22210267@viit.ac.in"]
            
            # Generate the MoM content from the to-do list
            client_groq_2= initialize_groq_client()
            todo_list_1= generate_todo_list(client_groq_2, meeting_summary)
            mom_content = generate_mom_format(todo_list_1)
            excel_file = mom_to_excel(mom_content, "MoM_Output.xlsx")  # Convert MoM text to Excel
            send_email_with_attachment(default_emails, "Meeting MoM", "Please find the MoM attached.", excel_file)


if __name__ == "__main__":
    main()