import streamlit as st
import spacy
import sqlite3
import pandas as pd
import requests
import msal
import os
import subprocess
from textblob import TextBlob
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder

# Ensure spaCy model is available before running
spacy_model = "en_core_web_sm"
try:
    nlp = spacy.load(spacy_model)
except OSError:
    subprocess.run(["python", "-m", "spacy", "download", spacy_model], check=True)
    nlp = spacy.load(spacy_model)

# Microsoft Outlook API Credentials (Using environment variables for security)
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
TENANT_ID = os.getenv("AZURE_TENANT_ID")

# Authenticate with Microsoft Graph API
def get_access_token():
    if not CLIENT_ID or not CLIENT_SECRET or not TENANT_ID:
        st.error("Error: Missing Azure credentials.")
        return None
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, 
        authority=f"https://login.microsoftonline.com/{TENANT_ID}", 
        client_credential=CLIENT_SECRET
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token['access_token'] if "access_token" in token else None

# Fetch Emails from Outlook (Run Only on Button Click)
def fetch_emails():
    token = get_access_token()
    if token:
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get("https://graph.microsoft.com/v1.0/me/messages", headers=headers)
        if response.status_code == 200:
            emails = response.json().get("value", [])
            for msg in emails:
                process_email(msg.get("subject", "No Subject"), msg.get("body", {}).get("content", ""))
            st.success("Emails Fetched Successfully!")
        else:
            st.error(f"Error fetching emails: {response.text}")

# Initialize SQLite database before use
def init_db():
    conn = sqlite3.connect("escalations.db")
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS escalations (
        id INTEGER PRIMARY KEY, 
        subject TEXT, 
        body TEXT, 
        status TEXT, 
        urgency TEXT, 
        entities TEXT
    )''')
    conn.commit()
    conn.close()

init_db()  # Ensure DB is initialized before Streamlit starts

# Process Email Content & Extract NLP Insights
def process_email(subject, body):
    sentiment_score = TextBlob(body).sentiment.polarity
    urgency = "High" if "urgent" in body.lower() or sentiment_score < -0.5 else "Normal"
    entities = [(ent.text, ent.label_) for ent in nlp(body).ents]
    log_to_database(subject, body, urgency, entities)

# Log data into database
def log_to_database(subject, body, urgency, entities):
    conn = sqlite3.connect("escalations.db")
    cursor = conn.cursor()
    cursor.execute("INSERT INTO escalations (subject, body, status, urgency, entities) VALUES (?, ?, ?, ?, ?)", 
                   (subject, body, "Pending", urgency, str(entities)))
    conn.commit()
    conn.close()

# AI Model for Predictive Insights (Fix Encoding for "urgency")
def train_escalation_model():
    try:
        data = pd.read_csv("escalations.csv")
        if data.empty:
            return None

        label_encoder = LabelEncoder()
        data["urgency_encoded"] = label_encoder.fit_transform(data["urgency"])

        X = data[['urgency_encoded']]
        y = data['status']
        model = RandomForestClassifier()
        model.fit(X, y)
        return model
    except Exception as e:
        st.error(f"Error training model: {e}")
        return None

escalation_model = train_escalation_model()

def predict_escalation_risk(urgency):
    if escalation_model:
        label_encoder = LabelEncoder()
        urgency_encoded = label_encoder.fit_transform([urgency])[0]
        return escalation_model.predict([[urgency_encoded]])[0]
    return "Unknown"

# Streamlit UI
st.title("EscalateAI - AI-powered Escalation Management")

if st.button("Fetch Emails"):
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
