from flask import Flask, request, jsonify
import json
from docx import Document
from datetime import datetime

app = Flask(__name__)

counter = 1

@app.route('/zap-webhook', methods=['POST'])
def webhook():
    data = request.json
    if data:
        global counter
        # Process the received data
        print(f'Received data: {data}')

        with open("webhook.json", "w") as file:
            json.dump(data, file, indent=4)
        
        json_data = extract_data(data["body"])
        personal_data = find_data(json_data[0])
        replace_data(personal_data)
        
        # Optionally, send a response back to Zapier
        response = {
            "status": "success",
            "message": "Data received successfully!"
        }

        counter += 1

        return jsonify(response), 200
    else:
        return jsonify({"status": "failure", "message": "No data received"}), 400

def extract_data(text):
    final_data = []
    if text:
        splited_texts = text.split("\n\n")
        first_name = splited_texts[0].split(": ")[1]
        last_name = splited_texts[1].split(": ")[1]
        address = splited_texts[2].split(": ")[1]
        city = splited_texts[3].split(": ")[1]
        state = splited_texts[4].split(": ")[1]
        zip = splited_texts[5].split(": ")[1]
        home_phone = splited_texts[6].split(": ")[1]
        cell_phone = splited_texts[7].split(": ")[1]
        email = splited_texts[8].split(": ")[1]
        agreed = splited_texts[9]
        emergency_name = splited_texts[10].split(": ")[1]
        emergency_phone = splited_texts[11].split(": ")[1]
        case_state = splited_texts[12].split(": ")[1]
        case_type = splited_texts[13].split(": ")[1]
        firm_name = splited_texts[14].split(": ")[1]
        attorney_name = splited_texts[15].split(": ")[1]
        attorney_phone = splited_texts[16].split(": ")[1]
        attorney_email = splited_texts[17].split(": ")[1]
        amount = splited_texts[18].split(": ")[1]
        payment_method = splited_texts[19].split(": ")[1]
        description = splited_texts[20].split(": ")[1]

        final_data.append({
            "first_name": first_name,
            "last_name": last_name,
            "address": address,
            "city": city,
            "state": state,
            "zip": zip,
            "home_phone": home_phone,
            "cell_phone": cell_phone,
            "email": email,
            "agreed": agreed,
            "emergency_name": emergency_name,
            "emergency_phone": emergency_phone,
            "case_state": case_state,
            "case_type": case_type,
            "firm_name": firm_name,
            "attorney_name": attorney_name,
            "attorney_phone": attorney_phone,
            "attorney_email": attorney_email,
            "amount": amount,
            "payment_method": payment_method,
            "description": description
        })
        return final_data
    else:
        return "Not found message"

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
    
    firm_address = json_data['attorney_email']
    print(f'<Lawyer_Email> ---> {firm_address}')

    loan_date = ""
    print(f'<Loan_Date> ---> {loan_date}')

    today = datetime.today()
    date_str = today.strftime("%Y-%m-%d")
    today_as_date = datetime.strptime(date_str, "%Y-%m-%d")
    case_id = today_as_date.strftime("%m%d%y") + str(f"{counter:02}")
    print(f'<Case_ID> ---> {case_id}')

    loan_amount = json_data['amount']
    print(f'<Loan_Amount> ---> {loan_amount}')

    handling_fee = "0"
    print(f'<Handling_Fee> ---> {handling_fee}')

    client_phone = json_data['home_phone']
    print(f'<Client_Phone> ---> {client_phone}')

    q1_interest = int(int(json_data['amount']) * 1.15)
    print(f'<Q1_Interest> ---> {q1_interest}')

    q2_interest = int(int(json_data['amount']) * 1.3)
    print(f'<Q2_Interest> ---> {q2_interest}')

    q3_interest = int(int(json_data['amount']) * 1.45)
    print(f'<Q3_Interest> ---> {q3_interest}')

    q4_interest = int(int(json_data['amount']) * 1.6)
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
        "<Q1_Interest>" : str(q1_interest),
        "<Q2_Interest>" : str(q2_interest),
        "<Q3_Interest>" : str(q3_interest),
        "<Q4_Interest>" : str(q4_interest),
        "<Client_Street>" : client_street,
        "<Client_CityState>" : client_citystate
    })

    return personal_data

def replace_data(personal_data):
    # Purchase Agreement.docx
    doc = Document('Purchase Agreement.docx')

    for key, value in personal_data[0].items():
        replace_word_in_paragraphs(doc.paragraphs, key, value)
        replace_word_in_tables(doc.tables, key, value)
    
    doc.save(f"{personal_data[0]["<Client_Name>"]}.docx")
    print(f"Words replaced and saved to {personal_data[0]["<Client_Name>"]}.docx")

if __name__ == '__main__':
    app.run(port=5002, debug=True)