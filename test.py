import streamlit as st
from groq import Groq
import speech_recognition as sr
import pyttsx3
import time
from datetime import datetime

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
    api_key = "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"  # Hardcoded API key
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

    Make sure each action item accurately reflects the content of the meeting summary and follows the format strictly.
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
            meeting_summary = st.text_area(
                "Enter your meeting summary:", 
                st.session_state.get('meeting_summary', ''),
                height=150
            )

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
                else:
                    st.error("Failed to generate the to-do list.")

if __name__ == "__main__":
    main()