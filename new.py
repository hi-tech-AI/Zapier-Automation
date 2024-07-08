from datetime import datetime

counter = 1

def generate_base_case_id(date):
    return date.strftime("%m%d%y")

def generate_case_id(current_date, last_counter):
    base_case_id = generate_base_case_id(current_date)
    next_number = last_counter + 1
    return f"{base_case_id}{next_number:02d}"

today = datetime.today()
case_id = generate_case_id(today, counter - 1)
print(f'<Case_ID> ---> {case_id}')