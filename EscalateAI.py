import streamlit as st
import spacy
import sqlite3
import pandas as pd
import requests
import msal
import os
from textblob import TextBlob
from celery import Celery
from trello import TrelloClient
from sklearn.ensemble import RandomForestClassifier

# Ensure spaCy model is available
spacy_model = "en_core_web_sm"
try:
    nlp = spacy.load(spacy_model)
except OSError:
    os.system(f"python -m spacy download {spacy_model}")
    nlp = spacy.load(spacy_model)

# Microsoft Outlook API Credentials (Using environment variables for security)
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
TENANT_ID = os.getenv("AZURE_TENANT_ID")

# Authenticate with Microsoft Graph API
def get_access_token():
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}", client_credential=CLIENT_SECRET)
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token['access_token'] if "access_token" in token else None

# Fetch Emails from Outlook
def fetch_emails():
    token = get_access_token()
    if token:
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get("https://graph.microsoft.com/v1.0/me/messages", headers=headers).json()
        for msg in response["value"]:
            process_email(msg["subject"], msg["body"]["content"])

# Initialize SQLite database
def init_db():
    conn = sqlite3.connect("escalations.db")
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS escalations (id INTEGER PRIMARY KEY, subject TEXT, body TEXT, status TEXT, urgency TEXT, entities TEXT)''')
    conn.commit()
    conn.close()

# Process Email Content
def process_email(subject, body):
    sentiment_score = TextBlob(body).sentiment.polarity
    urgency = "High" if "urgent" in body.lower() or sentiment_score < -0.5 else "Normal"
    entities = [(ent.text, ent.label_) for ent in nlp(body).ents]
    log_to_database(subject, body, urgency, entities)
    send_slack_notification(f"Escalation Logged: {subject}\nUrgency: {urgency}\nEntities: {entities}")
    create_trello_card(subject, body)
    if urgency == "High":
        escalate_case.apply_async((subject,), countdown=3600)

# Log data into database
def log_to_database(subject, body, urgency, entities):
    conn = sqlite3.connect("escalations.db")
    cursor = conn.cursor()
    cursor.execute("INSERT INTO escalations (subject, body, status, urgency, entities) VALUES (?, ?, ?, ?, ?)", (subject, body, "Pending", urgency, str(entities)))
    conn.commit()
    conn.close()

# Trello Integration
def create_trello_card(title, description):
    url = "https://api.trello.com/1/cards"
    query = {
        'name': title,
        'desc': description,
        'idList': os.getenv("TRELLO_LIST_ID"),
        'key': os.getenv("TRELLO_API_KEY"),
        'token': os.getenv("TRELLO_TOKEN")
    }
    requests.post(url, params=query)

# Slack Notifications
def send_slack_notification(message):
    webhook_url = os.getenv("SLACK_WEBHOOK_URL")
    requests.post(webhook_url, json={'text': message})

# Celery Task for Time-Based Escalation
celery = Celery('tasks', broker='redis://localhost:6379/0')
@celery.task
def escalate_case(subject):
    conn = sqlite3.connect("escalations.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM escalations WHERE subject=? AND status='Pending'", (subject,))
    if cursor.fetchone():
        send_slack_notification(f"⚠️ URGENT: Escalation not resolved: {subject}")
    conn.close()

# AI Model for Predictive Insights
def train_escalation_model():
    data = pd.read_csv("escalations.csv")
    X = data[['urgency']]
    y = data['status']
    model = RandomForestClassifier()
    model.fit(X, y)
    return model
escalation_model = train_escalation_model()

def predict_escalation_risk(urgency):
    return escalation_model.predict([[urgency]])[0]

# Streamlit UI
st.title("EscalateAI - AI-powered Escalation Management")
fetch_emails()
email_content = st.text_area("Paste Customer Email Here")
if st.button("Analyze Email"):
    doc = nlp(email_content)
    entities = [(ent.text, ent.label_) for ent in doc.ents]
    log_to_database("Manual Entry", email_content, "Normal", entities)
    st.success("Escalation Logged Successfully!")
    st.write("Entities Identified:", entities)

st.subheader("Past Escalations")
st.dataframe(pd.read_sql_query("SELECT * FROM escalations", sqlite3.connect("escalations.db")))

# Initialize database on first run
init_db()
