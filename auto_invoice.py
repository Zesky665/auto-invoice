import gspread
import datetime as dt
import numpy as np

WAGE = 22

def extract_num(str):
    """extract_num("NO.X") -> X"""
    num = [int(m) for m in str.split(".") if m.isdigit()]
    if num[0]:
        return num[0]
    else:
        print("No num found")

def update_invoice_number(sheet, cell):
    """update_invoice_number(invoice_num) -> new_invoice_num"""
    # Update Invoice Number
    current_number = worksheet.acell(cell).value
    invoice_number = extract_num(current_number)
    invoice_number += 1
    sheet.update(cell, f'NO.{invoice_number}')

def update_send_date(sheet, now, cell):
    """update_send_date(sheet) -> sheet"""
    send_date = now.strftime("%d.%m.%Y")
    sheet.update(cell, f'DATE:{send_date}')

def find_start_date():
    """find_start_date(now) -> XX.MM.YYYY"""
    st_d = now - dt.timedelta(days=30)
    st_d = st_d.replace(day=16)
    if st_d.strftime("%a") == "Sat":
        st_d = st_d + dt.timedelta(days=2)
    elif st_d.strftime("%a") == "Sun":
        st_d = st_d + dt.timedelta(days=1)
    return st_d

def find_end_date():
    """find_start_date(now) -> XX.MM.YYYY"""
    st_d = now
    st_d = st_d.replace(day=15)
    if st_d.strftime("%a") == "Sat":
        st_d = st_d - dt.timedelta(days=1)
    elif st_d.strftime("%a") == "Sun":
        st_d = st_d - dt.timedelta(days=2)
    return st_d

def get_working_hours(start_date, end_date):
    """get_working_hours(start_date, end_date) -> xyz"""
    start = start_date.strftime("%Y-%m-%d")
    end = end_date.strftime("%Y-%m-%d")
    working_days = np.busday_count(start, end, weekmask=[1,1,1,1,1,0,0])
    return working_days * 8

def generate_billable_hours_line_text(start_date, end_date):
    """generate_billable_hours_line_text(start_date, end_date) -> text"""

    default_hours = get_working_hours(s_d, e_d)
    start = start_date.strftime("%d.%m.%Y")
    end = end_date.strftime("%d.%m.%Y")

    return f'''1.Quality Assureance Development {start} to {end} ({default_hours}h, fee {WAGE} € per hour )'''

def add_line_item(sheet, row, line, amount):
    """ add_line_item(sheet, row, line) -> sheet """
    sheet.update(f'A{row}', line)
    sheet.update(f'J{row}', int(amount))

service_account = gspread.service_account(filename="auto-invoice-service-account-file.json")

sheet = service_account.open("template")

worksheet = sheet.worksheet("Sheet1")

# Update invoice number
update_invoice_number(worksheet, 'H5')

# Update the Send Date
now = dt.datetime.today()

update_send_date(worksheet, now, "I11")

# Add Billable hours
s_d = find_start_date()
e_d = find_end_date()

line = generate_billable_hours_line_text(s_d, e_d)
amount = get_working_hours(s_d, e_d) * WAGE
add_line_item(worksheet, 20, line, amount)
