from flask import Flask, request, render_template, send_file
import pandas as pd
import numpy as np
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        # Get the uploaded file
        file = request.files['file']
        if not file:
            return "Error: No file uploaded."

        # Read weights, impacts, and email
        weights = request.form['weights']
        impacts = request.form['impacts']
        email = request.form['email']

        # Validate inputs
        weights = list(map(float, weights.split(',')))
        impacts = impacts.split(',')

        if len(weights) != len(impacts):
            return "Error: The number of weights and impacts must be equal."
        if any(i not in ['+', '-'] for i in impacts):
            return "Error: Impacts must be '+' or '-'."
        if '@' not in email:
            return "Error: Invalid email ID."

        # Save the uploaded file temporarily
        input_file_path = 'uploaded_file.xlsx'
        file.save(input_file_path)

        # Perform TOPSIS calculation
        result_file_path = 'topsis_result.csv'
        perform_topsis(input_file_path, weights, impacts, result_file_path)

        # Send the result file via email
        send_email(email, result_file_path)

        # Clean up temporary files
        os.remove(input_file_path)
        os.remove(result_file_path)

        return "TOPSIS calculation completed. Results have been sent to your email."

    except Exception as e:
        return f"An error occurred: {e}"

def perform_topsis(input_file, weights, impacts, output_file):
    df = pd.read_excel(input_file)
    cols = df.columns[1:]

    # Normalize the data
    root_sum_sq = np.sqrt(np.sum(df[cols] ** 2, axis=0))
    df[cols] = df[cols] / root_sum_sq

    # Apply weights
    normalized_weights = weights / np.sum(weights)
    df[cols] = df[cols] * normalized_weights

    # Determine best and worst values
    best_values = df[cols].max(axis=0)
    worst_values = df[cols].min(axis=0)

    for i, impact in enumerate(impacts):
        if impact == '-':
            best_values[i], worst_values[i] = worst_values[i], best_values[i]

    # Calculate distances
    df['best_dist'] = np.sqrt(np.sum((df[cols] - best_values) ** 2, axis=1))
    df['worst_dist'] = np.sqrt(np.sum((df[cols] - worst_values) ** 2, axis=1))

    # Calculate Topsis score and ranks
    df['Topsis Score'] = df['worst_dist'] / (df['best_dist'] + df['worst_dist'])
    df['Rank'] = df['Topsis Score'].rank(ascending=False).astype(int)

    df.to_csv(output_file, index=False)

def send_email(recipient, file_path):
    sender_email = "samplemashup12@gmail.com"
    sender_password = "vcgdcnxfxqiufmgv"

    # Email content
    subject = "TOPSIS Results"
    body = "Please find attached the TOPSIS result file."

    # Setup email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Attach file
    with open(file_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(file_path)}")
        msg.attach(part)

    # Send email
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)

if __name__ == '__main__':
    app.run(debug=True)
