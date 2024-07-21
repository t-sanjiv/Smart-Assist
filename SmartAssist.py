import tkinter as tk
from tkinter import scrolledtext, messagebox, Toplevel
import pandas as pd
from transformers import pipeline
import random
from datetime import datetime
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression

# Initialize the table question answering pipeline
t2 = pipeline(task="table-question-answering", model="neulab/omnitab-large-finetuned-wtq")

# Global variables to store phone number, OTP, and tables
phone_number = None
otp_verified = False
issue_description_requested = False
ticket_created = False
table_customer = None
table_log = None
table_ticket = None
DeviceID = None  
table_alert = None

# Define greeting inputs and responses
GREET_INPUTS = ("hello", "hi", "greetings", "sup", "what's up", "hey")
GREET_RESPONSES = ["Hi! I'm SmartAssist. Could you please provide your phone number in which you're facing an issue?"]

def greet(sentence):
    for word in sentence.lower().split():
        if word.lower() in GREET_INPUTS:
            return random.choice(GREET_RESPONSES)

# Function to load tables and handle OTP verification and data retrieval
def process_data(user_input):
    global phone_number, otp_verified, issue_description_requested, ticket_created, table_customer, table_log, table_ticket, DeviceID, table_alert

    if not phone_number:
        # If phone number is not entered yet, treat it as phone number input
        phone_number = user_input.strip().lower()
        text_area.insert(tk.END, f"You: {user_input}\n\n")
        text_area.insert(tk.END, "Support Engineer: Please enter OTP to verify.\n\n")
        user_entry.delete(0, tk.END)

        # Load customer details from Excel
        excel_file_customer = "D:\\Datas\\CustomerDB.xlsx"
        sheet_name_customer = "Sheet1"
        xls_customer = pd.ExcelFile(excel_file_customer)
        table_customer = pd.read_excel(xls_customer, sheet_name=sheet_name_customer)
        table_customer = table_customer.astype(str)

        # Load log details from Excel
        excel_file_log = "D:\\Datas\\LogDetails.xlsx"
        sheet_name_log = "Sheet1"
        xls_log = pd.ExcelFile(excel_file_log)
        table_log = pd.read_excel(xls_log, sheet_name=sheet_name_log)
        table_log = table_log.astype(str)

        # Load ticketing system details from Excel
        excel_file_ticket = "D:\\Datas\\TicketingSystem.xlsx"
        sheet_name_ticket = "Sheet1"
        xls_ticket = pd.ExcelFile(excel_file_ticket)
        table_ticket = pd.read_excel(xls_ticket, sheet_name=sheet_name_ticket)
        table_ticket = table_ticket.astype(str)

        # Load alert log details from Excel
        excel_file_alert = "D:\\Datas\\Alertlog.xlsx"
        sheet_name_alert = "Sheet1"
        xls_alert = pd.ExcelFile(excel_file_alert)
        table_alert = pd.read_excel(xls_alert, sheet_name=sheet_name_alert)
        table_alert = table_alert.astype(str)

        # Queries to fetch customer details
        query1_customer = "Give me the IMEI corresponding to " + phone_number
        query2_customer = "Give me the Customer corresponding to " + phone_number
        query3_customer = "Give me the Address corresponding to " + phone_number
        query4_customer = "Give me the Email corresponding to " + phone_number

        # Fetch customer details
        IMEI = t2(table=table_customer, query=query1_customer)["answer"]
        Name = t2(table=table_customer, query=query2_customer)["answer"]
        Address = t2(table=table_customer, query=query3_customer)["answer"]
        Email = t2(table=table_customer, query=query4_customer)["answer"]

        # Queries to fetch log details
        query1_log = "Give me the Date and Time corresponding to " + IMEI
        query2_log = "Give me the Type corresponding to " + IMEI
        query3_log = "Give me the Network Type corresponding to " + IMEI
        query4_log = "Give me the DeviceID corresponding to " + IMEI

        # Fetch log details
        DeviceID = t2(table=table_log, query=query4_log)["answer"]
        Timestamp = t2(table=table_log, query=query1_log)["answer"]
        Type = t2(table=table_log, query=query2_log)["answer"]
        Network = t2(table=table_log, query=query3_log)["answer"]

        # Search for DeviceID in ticketing system
        word_found = False
        for col in table_ticket.columns:
            for value in table_ticket[col]:
                if isinstance(value, str) and DeviceID in value:
                    word_found = True
                    break
            if word_found:
                break

        if word_found:
            query1_ticket = "Give me the Title related to " + DeviceID
            query2_ticket = "Give me the TicketID related to " + DeviceID
            query3_ticket = "Give me the Status related to " + DeviceID
            query4_ticket = "Give me the Date and Time of creation related to " + DeviceID

            # Fetch ticketing details
            Title = t2(table=table_ticket, query=query1_ticket)["answer"]
            TicketID = t2(table=table_ticket, query=query2_ticket)["answer"]
            Status = t2(table=table_ticket, query=query3_ticket)["answer"]
            TicketTimestamp = t2(table=table_ticket, query=query4_ticket)["answer"]

        else:
            # If no ticket found, set default values
            Title = "No ticket found. Ticket to be created"
            TicketID = ""
            Status = ""
            TicketTimestamp = ""

        # Display output in a separate window
        display_data_window(IMEI, Name, Address, Email, Timestamp, Type, Network, DeviceID, Title, TicketID, Status, TicketTimestamp)

        # Clear the user input field
        user_entry.delete(0, tk.END)

    elif not otp_verified:
        # OTP verification logic (for demo purpose, just checking if user entered OTP)
        otp = user_input.strip().lower()
        if not otp:
            messagebox.showerror("Error", "Invalid OTP")
            return
        otp_verified = True
        process_data(user_input)  # Re-trigger data processing after OTP verification

    elif not issue_description_requested:
        # Request issue description from the customer
        text_area.insert(tk.END, f"You: {user_input}\n\n")
        text_area.insert(tk.END,"Support Engineer: Please enter the issue are you facing.\n\n")
        issue_description_requested = True
        user_entry.delete(0, tk.END)

    elif not ticket_created:
        issue_description = user_input.strip()
        text_area.insert(tk.END, f"You: {user_input}\n\n")
        # Check if a ticket exists for the DeviceID
        word_found = False
        for col in table_ticket.columns:
            for value in table_ticket[col]:
                if isinstance(value, str) and DeviceID in value:
                    word_found = True
                    break
            if word_found:
                break

        if word_found:
            text_area.insert(tk.END, "Support Engineer: This issue has already been raised and will be resolved soon.\n\n")
        else:
            # Check if DeviceID exists in table_alert
            if str(int(DeviceID)) in table_alert['DeviceID'].values:
                # DeviceID found in table_alert, proceed to create a new ticket
                alert_row = table_alert[table_alert['DeviceID'] == str(int(DeviceID))].iloc[0]
                # Prepare data for the new ticket
                next_ticket_id = int(table_ticket['TicketID'].max()) + 1
                title = alert_row['Summary']
                description = ' '.join(alert_row.dropna().astype(str))
                severity = ''  # Determine how severity is derived
                status = 'Open'
                created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Current timestamp
                closed_at = ''  # Initially empty

                # Create a new ticket record
                new_ticket = {
                    'TicketID': next_ticket_id,
                    'Title': title,
                    'Description': description,
                    'Severity': severity,
                    'Status': status,
                    'CreatedAt': created_at,
                    'ClosedAt': closed_at
                }

                # Convert new_ticket to DataFrame and append to table_ticket
                df_new_ticket = pd.DataFrame([new_ticket])
                table_ticket = pd.concat([table_ticket, df_new_ticket], ignore_index=True)

                # Save updated TicketingSystem back to Excel
                table_ticket.to_excel('D:\\Datas\\TicketingSystem.xlsx', index=False)

                # Respond to the customer
                text_area.insert(tk.END, "Support Engineer: Your issue has been noted and a ticket has been created.\n\n")
            else:
                # DeviceID not found in table_alert
                # Prepare data for the new ticket
                next_ticket_id = int(table_ticket['TicketID'].max()) + 1
                title = issue_description  # Use issue description as title
                description = issue_description  # Use issue description as description
                severity = ''  # Determine how severity is derived
                status = 'Open'
                created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Current timestamp
                closed_at = ''  # Initially empty

                # Create a new ticket record
                new_ticket = {
                'TicketID': next_ticket_id,
                'Title': title,
                'Description': description,
                'Severity': severity,
                'Status': status,
                'CreatedAt': created_at,
                'ClosedAt': closed_at
                }

                # Convert new_ticket to DataFrame and append to table_ticket
                df_new_ticket = pd.DataFrame([new_ticket])
                table_ticket = pd.concat([table_ticket, df_new_ticket], ignore_index=True)

                # Save updated TicketingSystem back to Excel
                table_ticket.to_excel('D:\\Datas\\TicketingSystem.xlsx', index=False)

                # Respond to the customer
                text_area.insert(tk.END, "Support Engineer: Your issue has been noted and a ticket has been created.\n\n")

        ticket_created = True
        user_entry.delete(0, tk.END)

    elif user_input.lower() == 'bye':
        # If user says bye, exit the conversation
        text_area.insert(tk.END, f"You: {user_input}\n\n")
        text_area.insert(tk.END, "Goodbye! Take care.\n\n")
        close_windows()  

    # Clear the user input field
    user_entry.delete(0, tk.END)

def close_windows():
    # Close both windows and restart chatbot after 2 seconds
    global root, data_window
    if data_window:
        data_window.destroy()
    root.after(2000, reset_chatbot_state)

def reset_chatbot_state():
    global phone_number, otp_verified, issue_description_requested, ticket_created
    phone_number = None
    otp_verified = False
    issue_description_requested = False
    ticket_created = False
    text_area.delete(1.0, tk.END)  
    text_area.insert(tk.END, greet("hello") + "\n\n")

def display_data_window(
    imei, name, address, email, timestamp, type_val, network, deviceid, title, ticketid, status, ticket_timestamp):
    global data_window
    data_window = Toplevel(root)
    data_window.title("SmartAssist Data Display")
    
    # Create labels and display data
    label_customer = tk.Label(data_window, text="Customer Details:")
    label_customer.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_imei = tk.Label(data_window, text=f"IMEI: {imei}")
    label_imei.grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_name = tk.Label(data_window, text=f"Name: {name}")
    label_name.grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_address = tk.Label(data_window, text=f"Address: {address}")
    label_address.grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_email = tk.Label(data_window, text=f"Email: {email}\n")
    label_email.grid(row=4, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_log = tk.Label(data_window, text="Log Details:")
    label_log.grid(row=5, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_timestamp = tk.Label(data_window, text=f"Timestamp: {timestamp}")
    label_timestamp.grid(row=6, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_type = tk.Label(data_window, text=f"Type: {type_val}")
    label_type.grid(row=7, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_network = tk.Label(data_window, text=f"Network: {network}")
    label_network.grid(row=8, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_deviceid = tk.Label(data_window, text=f"DeviceID: {deviceid}\n")
    label_deviceid.grid(row=9, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_ticketing = tk.Label(data_window, text="Ticketing Details:")
    label_ticketing.grid(row=10, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_title = tk.Label(data_window, text=f"Title: {title}")
    label_title.grid(row=11, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_ticketid = tk.Label(data_window, text=f"TicketID: {ticketid}")
    label_ticketid.grid(row=12, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_status = tk.Label(data_window, text=f"Status: {status}")
    label_status.grid(row=13, column=0, padx=10, pady=5, sticky=tk.W)
    
    label_ticket_timestamp = tk.Label(data_window, text=f"Ticket Creation: {ticket_timestamp}")
    label_ticket_timestamp.grid(row=14, column=0, padx=10, pady=5, sticky=tk.W)


# Create the main application window
root = tk.Tk()
root.title("SmartAssist Chatbot")

# Frame to hold chat interface components
chat_frame = tk.Frame(root)
chat_frame.pack(padx=20, pady=20)

# ScrolledText widget to display conversation
text_area = scrolledtext.ScrolledText(chat_frame, width=80, height=20)
text_area.pack(fill=tk.BOTH, expand=True)

# Entry widget for user input
user_entry = tk.Entry(chat_frame, width=60)
user_entry.pack(side=tk.LEFT, padx=10, pady=10)

# Button to submit user input
submit_button = tk.Button(chat_frame, text="Send", command=lambda: process_data(user_entry.get()))
submit_button.pack(side=tk.RIGHT, padx=10, pady=10)

# Start the conversation with a greeting and prompt for phone number
text_area.insert(tk.END, greet("hello") + "\n\n")

# Initialize global variable for data window
data_window = None

# Start the main event loop
root.mainloop()

# Check if Level 3 (Field Engineer)is required
# Step 1: Read data from Excel (NewFieldEngineer.xlsx)
excel_file = 'D:\\Datas\\FieldEngineerReq.xlsx'  
df = pd.read_excel(excel_file)

# Step 2: Verify data structure (column names)
# print(df.columns)  # Check column names to ensure 'Situation' and 'Label' are correctly named

# Step 3: Preprocess data (assuming 'Situation' and 'Label' columns)
X = df['Situation '].astype(str)
y = df['Label']

# Step 4: Feature extraction using TF-IDF vectorization
vectorizer = TfidfVectorizer(max_features=1000)  # Adjust max_features as needed
X_vec = vectorizer.fit_transform(X)

# Step 5: Train a logistic regression model
model = LogisticRegression(max_iter=1000)
model.fit(X_vec, y)

# Step 6: Read data from another Excel file (TicketingSystem.xlsx)
ticket_excel_file = 'D:\\Datas\\TicketingSystem.xlsx'  
ticket_df = pd.read_excel(ticket_excel_file)

# Step 7: Use 'Description' column for prediction
descriptions = ticket_df['Description'].astype(str)
descriptions_vec = vectorizer.transform(descriptions)
predictions = model.predict(descriptions_vec)

# Step 8: Create a new window using Tkinter
root = tk.Tk()
root.title("Level 3 Escalation Requirement")
txt = scrolledtext.ScrolledText(root, width=150, height=30)
txt.grid(column=0, row=0, padx=10, pady=10)

# Display predictions in the scrolled text widget
for description, prediction in zip(descriptions, predictions):
    txt.insert(tk.END, f"{description}: {prediction}\n\n")

# Start the Tkinter main loop
root.mainloop()
