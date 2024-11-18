import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import streamlit as st

def send_email(to_emails, subject, body):
    from_email = "harsh.22210267@viit.ac.in"  
    password = "noxp vggx qcyu exfs"  

    for email in to_emails:
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        try:
            # Print the status of the connection steps
            st.info(f"Connecting to SMTP server...")
            server = smtplib.SMTP('smtp.gmail.com', 587)
            st.info(f"Connection successful. Starting TLS...")
            server.starttls()
            st.info(f"Logging in with email: {from_email}...")
            server.login(from_email, password)
            st.success(f"Logged in successfully.")

            # Print the email sending status
            text = msg.as_string()
            st.info(f"Sending email to {email}...")
            server.sendmail(from_email, email, text)
            server.quit()

            st.success(f"Email successfully sent to {email}")
        except smtplib.SMTPAuthenticationError:
            st.error("Authentication error: Invalid email or password. Make sure you're using an App Password.")
        except smtplib.SMTPException as e:
            st.error(f"SMTP error: {e}")
        except Exception as e:
            st.error(f"An unexpected error occurred: {e}")

# Example usage within Streamlit
if st.button("📧 Send To-Do List via Email"):
    default_emails = ["amisha.22211424@viit.ac.in", "mitesh.22210204@viit.ac.in", "hrushikesh.22210850@viit.ac.in"]
    send_email(default_emails, "Meeting To-Do List", "Here is the to-do list from the meeting.")



