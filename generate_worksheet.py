import time

import gspread
import pandas as pd
from pandas import DataFrame

import main as main_script
from salesforce import SalesForce
import numpy as np


def main():
    gc = gspread.service_account(filename='credentials.json')

    while True:
        try:
            sh = gc.open_by_key(main_script.SPREADSHEET_KEY)
            break
        except Exception as e:
            print(e)
            print("trying again")
            time.sleep(2)
            pass
    worksheet = sh.worksheet("Completed Master List")

    df: DataFrame = main_script.get_df_from_sheet(worksheet)  # fix header
    df.replace("", np.nan, inplace=True)  # switch empty strings with nan to use dropna method
    df.dropna(how="all", inplace=True)  # delete empty rows

    # check that every row has a company name and DOt
    for index, row in df.iterrows():
        company_name = row[main_script.COMPANY_COLUMN_NAME]
        dot_number = row[main_script.DOT_COLUMN_NAME]
        if dot_number == "":
            raise ValueError(f"\"{company_name}\" does not have a DOT number")

    df.replace(np.nan, "", inplace=True)  # switch NaN back to empty string

    all_companies = df[main_script.COMPANY_COLUMN_NAME].to_list()
    all_dot_numbers = df[main_script.DOT_COLUMN_NAME].to_list()
    all_data = [(company, dot_number) for company, dot_number in zip(all_companies, all_dot_numbers)]
    all_data = list(set(all_data))
    all_companies = [company for company, dot in all_data]
    all_dot_numbers = [dot for company, dot in all_data]
    blanks = ["" for _ in all_companies]
    excel_data = {
        main_script.DOT_COLUMN_NAME: all_dot_numbers,
        main_script.COMPANY_COLUMN_NAME: all_companies,
        main_script.EMAIL_KEY: blanks,
        main_script.COMPLETION_STATUS_KEY: blanks
    }

    new_df = pd.DataFrame(excel_data, dtype=str)
    new_df.sort_values(by=main_script.COMPANY_COLUMN_NAME, ascending=True, inplace=True)
    new_df.reset_index(drop=True, inplace=True)

    sf = SalesForce(main_script.SF_USERNAME, main_script.SF_PASSWORD)
    print("Logging into Salesforce...")
    sf.login()
    print("Getting emails from Salesforce...")
    num_rows = new_df.shape[0]
    for index, row in new_df.iterrows():
        print(f"Email addresses scraped: {int(index) + 1} / {num_rows}")
        dot_number = row[main_script.DOT_COLUMN_NAME]
        email_address = sf.get_email(dot_number)
        row[main_script.EMAIL_KEY] = email_address
    new_df.to_excel("Companies.xlsx", index=False)


if __name__ == "__main__":
    main()
