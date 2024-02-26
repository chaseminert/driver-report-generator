import os
import signal
import time

import gspread
import pandas as pd
import win32com.client as win32
from pandas.core.frame import DataFrame

import generate_worksheet

SF_USERNAME = "Salesforce username here"
SF_PASSWORD = "Salesforce password here"
SPREADSHEET_KEY = "Spreadsheet key here"

DOT_COLUMN_NAME = "DOT #"  # key for Excel and google sheets
COMPANY_COLUMN_NAME = "Company Name"  # key for Excel and google sheets
DRIVER_FIRST_NAME = "First Name"
DRIVER_LAST_NAME = "Last Name"
NOT_FOUND = "NOT FOUND"


COMPLETE_STATUS = "Complete"
ERROR_STATUS = "Error"

COMPLETION_STATUS_KEY = "Completion Status"  # key for Excel
EMAIL_KEY = "Email"
NO_EMAIL_VALUE = "NO EMAIL"

excel_path = "Companies.xlsx"

exit_requested = False
started_sending = False

testing = False  # just save the emails


def main():
    print("PRESS CTRL+C TO EXIT\n")

    global exit_requested
    global started_sending
    global testing
    signal.signal(signal.SIGINT, signal_handler)

    if not os.path.exists(excel_path):
        print("Getting company list...")
        generate_worksheet.main()
        print(f"\nCheck \"{excel_path}\" and fill in the missing email addresses.\nThen, re-run the program.\n")
        exit()

    excel_df = pd.read_excel(excel_path, dtype=str)

    print("Scraping google sheets...")
    google_sheets_df = get_google_sheets_df()

    print("Sending emails...\n")
    started_sending = True
    num_rows = excel_df.shape[0]
    for index, row in excel_df.iterrows():
        status = row[COMPLETION_STATUS_KEY]
        email = row[EMAIL_KEY]
        if status == COMPLETE_STATUS:
            continue
        if email == NO_EMAIL_VALUE:
            row[COMPLETION_STATUS_KEY] = COMPLETE_STATUS
            continue
        company_name = row[COMPANY_COLUMN_NAME]
        dot_number = row[DOT_COLUMN_NAME]
        actual_index = int(index) + 1
        send_email_per_company(company_name, dot_number, email, google_sheets_df, actual_index)
        row[COMPLETION_STATUS_KEY] = COMPLETE_STATUS
        if not testing:
            pass
            # save_excel_file(excel_df)
        print(f"{company_name} is complete ({actual_index} / {num_rows})")
        if exit_requested:
            exit()


def send_email_per_company(company, dot_number, email_address, google_sheets_df, index=0):
    # all of this df stuff just fixes the header

    drivers = get_drivers(dot_number, google_sheets_df)
    send_outlook_email(email_address, company, drivers, index=index)

    # print("Drivers:\n")
    # for driver in drivers:
    #     print(driver)
    # print(f"\nEmail Address: {email_address}")


def get_df_from_sheet(worksheet):
    df = pd.DataFrame(worksheet.get_all_values(), dtype=str)
    new_header = df.iloc[0]  # Grab the first row for the header
    df = df[1:]  # Take the data less the header row
    df.columns = new_header  # Set the header row as the df header
    df.reset_index(drop=True, inplace=True)
    return df


def get_dot_number(company_name_target: str, df: DataFrame):
    for index, row in df.iterrows():
        company_name = row[COMPANY_COLUMN_NAME]
        if company_name == company_name_target:
            dot_number = row[DOT_COLUMN_NAME]
            return dot_number
    return NOT_FOUND


def get_drivers(dot_number_target: str, df: DataFrame):
    drivers = []
    for index, row in df.iterrows():
        dot_number = row[DOT_COLUMN_NAME]
        if dot_number == dot_number_target:
            driver_first_name = row[DRIVER_FIRST_NAME].strip()
            driver_last_name = row[DRIVER_LAST_NAME].strip()
            driver_full_name = f"{driver_first_name} {driver_last_name}"
            if driver_full_name == " ":
                raise ValueError(f"DOT Number: {dot_number} has a blank driver")
            drivers.append(driver_full_name)
    return drivers


# Doesn't actually send the email yet
def send_outlook_email(email, company_name, drivers: list, index):
    global testing
    driver_placeholder = "<p class=MsoNormal>[Insert List of CDL Drivers w/ Bullet points]<o:p></o:p>"
    company_placeholder = "[Company Name]"
    driver_list_str = get_driver_list_str(drivers)

    current_dir = os.getcwd()
    email_save_path = os.path.join(current_dir, "Emails", f"Email {index}.msg")
    template_path = os.path.join(current_dir, "Template.oft")

    outlook = win32.Dispatch('outlook.application')  # Start Outlook application
    mail = outlook.CreateItemFromTemplate(template_path)
    mail.SentOnBehalfOfName = "drugreporting@transcompservice.com"
    mail_body = mail.HTMLBody.replace(driver_placeholder, driver_list_str).replace(company_placeholder, company_name)
    mail.HTMLBody = mail_body
    mail.To = email
    # mail.Send()  # FOR TESTING

    if testing:
        mail.SaveAs(email_save_path)
    else:
        mail.Send()


def get_email_body(company_name, drivers: list):
    driver_list_str = get_driver_list_str(drivers)
    body = f"Hello {company_name}, we have the following drivers listed as active:\n\n"
    body += driver_list_str
    body += "\nIf anyone is inactive, please let us know. Thank you!"
    return body


def save_excel_file(df):
    global excel_path
    global exit_requested
    try:
        df.to_excel(excel_path, index=False)
    except KeyboardInterrupt:
        print("saving before exit...")
        df.to_excel(excel_path, index=False)
        print("saved")
        exit()


def get_google_sheets_df():
    gc = gspread.service_account(filename='credentials.json')

    while True:
        try:
            sh = gc.open_by_key(SPREADSHEET_KEY)
            break
        except Exception as e:
            print(e)
            time.sleep(2)
            pass
    worksheet = sh.worksheet("Completed Master List")
    google_sheets_df = get_df_from_sheet(worksheet)
    return google_sheets_df


def get_driver_list_str(drivers: list):
    return "<ul>" + "".join([f"<li>{driver}</li>" for driver in drivers]) + "</ul>"


def signal_handler():
    global started_sending
    global exit_requested
    if not started_sending:
        exit()
    exit_requested = True
    print("Saving before exit...")


if __name__ == "__main__":
    main()
