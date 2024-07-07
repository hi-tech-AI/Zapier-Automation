from flask import Flask, request, jsonify
import json
from docx import Document
from docx.shared import Inches
import os
from dotenv import load_dotenv

load_dotenv()

document_name = os.getenv("DOCUMENT_NAME")

app = Flask(__name__)

@app.route('/zap-webhook', methods=['POST'])
def webhook():
    data = request.json
    if data:
        # Process the received data
        print(f'Received data: {data}')

        with open("webhook.json", "w") as file:
            json.dump(data, file, indent=4)

        personal_data = find_data(data)
        replace_data(personal_data)
        
        # Optionally, send a response back to Zapier
        response = {
            "status": "success",
            "message": "Data received successfully!"
        }
        return jsonify(response), 200
    else:
        return jsonify({"status": "failure", "message": "No data received"}), 400

def replace_word_in_paragraphs(paragraphs, old_word, new_word):
    for para in paragraphs:
        if old_word in para.text:
            para.text = para.text.replace(old_word, new_word)

def replace_word_in_tables(tables, old_word, new_word):
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                replace_word_in_paragraphs(cell.paragraphs, old_word, new_word)

def find_data(json_data):
    personal_data = []

    client_name = json_data['first_name'] + ' ' + json_data['last_name']
    print(f'<Client_Name> ---> {client_name}')

    client_email = json_data['email']
    print(f'<Client_Email> ---> {client_email}')

    firm_name = json_data['firm_name']
    print(f'<Firm_Name> ---> {firm_name}')
    
    attorney_name = json_data['attorney_name']
    print(f'<Lawyer_Name> ---> {attorney_name}')
    
    firm_address = ""
    print(f'<Lawyer_Email> ---> {firm_address}')

    loan_date = ""
    print(f'<Loan_Date> ---> {loan_date}')

    case_id = ""
    print(f'<Case_ID> ---> {case_id}')

    loan_amount = json_data['amount']
    print(f'<Loan_Amount> ---> {loan_amount}')

    handling_fee = ""
    print(f'<Handling_Fee> ---> {handling_fee}')

    client_phone = json_data['home_phone']
    print(f'<Client_Phone> ---> {client_phone}')

    q1_interest = ""
    print(f'<Q1_Interest> ---> {q1_interest}')

    q2_interest = ""
    print(f'<Q2_Interest> ---> {q2_interest}')

    q3_interest = ""
    print(f'<Q3_Interest> ---> {q3_interest}')

    q4_interest = ""
    print(f'<Q4_Interest> ---> {q4_interest}')

    client_street = json_data['address']
    print(f'<Client_Street> ---> {client_street}')

    client_citystate = json_data['city'] + ", " + json_data['state'] + " " + json_data['zip']
    print(f'<Client_CityState> ---> {client_citystate}')

    personal_data.append({
        "<Client_Name>" : client_name,
        "<Client_Email>" : client_email,
        "<Firm_Name>" : firm_name,
        "<Lawyer_Name>" : attorney_name,
        "<Lawyer_Email>" : firm_address,
        "<Loan_Date>" : loan_date,
        "<Case_ID>" : case_id,
        "<Loan_Amount>" : loan_amount,
        "<Handling_Fee>" : handling_fee,
        "<Client_Phone>" : client_phone,
        "<Q1_Interest>" : q1_interest,
        "<Q2_Interest>" : q2_interest,
        "<Q3_Interest>" : q3_interest,
        "<Q4_Interest>" : q4_interest,
        "<Client_Street>" : client_street,
        "<Client_CityState>" : client_citystate
    })

    return personal_data

def replace_data(personal_data):
    # Purchase Agreement.docx
    doc = Document(document_name)

    for key, value in personal_data[0].items():
        replace_word_in_paragraphs(doc.paragraphs, key, value)
        replace_word_in_tables(doc.tables, key, value)
        print(f"Replace {key} to {value}")

    doc.save(f"{personal_data[0]["<Client_Name>"]}.docx")
    print(f"Words replaced and saved to {personal_data[0]["<Client_Name>"]}.docx")

if __name__ == '__main__':
    app.run(port=5002, debug=True)