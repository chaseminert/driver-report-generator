# Driver Report Generator

## Overview
The Driver Report Generator is an internal tool developed for Transportation Compliance Service to streamline the process of updating the drug testing consortium. This project automates the generation and delivery of driver reports to clients, enhancing operational efficiency and client satisfaction.

## Key Features
- **Real-Time Data Scraping**: Scrapes data from a Google Spreadsheet containing driver consortium information in real time.
- **Dynamic Email Generation**: Dynamically generates emails listing active drivers for each company and sends them to the respective company contacts.
- **Salesforce Integration**: Utilizes the Salesforce API to retrieve email addresses for company contacts, ensuring accurate and up-to-date contact information.
- **Automated Email Sending**: Automates the process of sending emails using Outlook templates, reducing manual effort and improving efficiency.

## Technologies Used
- **Python**: Core programming language for developing the application logic.
- **Google Sheets API**: Used for real-time data scraping from the Google Spreadsheet.
- **Pandas**: Used for processing the sheet data
- **Salesforce API**: Integrated to retrieve email addresses for company contacts.
- **Win32**: Library for sending emails programmatically using Outlook.

## Project Specifics
The Driver Report Generator project is open source but intended for internal use by Transportation Compliance Service. While the source code is publicly available for review, it is tailored to fit TCS's specific needs and may not be suitable for external use. Contributions from external parties outside of Transportation Compliance Service are not accepted.

## Challenges Faced and Learning Experience
### Real-Time Data Scraping
One significant challenge was figuring out how to scrape data from the Google Spreadsheet in real time. Through reading the documentation and exploring available resources, I learned about the Google Sheets API and service accounts. Additionally, I discovered a helpful wrapper for the Google Sheets API that facilitated the conversion of spreadsheet data into a Pandas DataFrame, streamlining the data processing workflow.

### Accessing Company Emails
Another challenge was obtaining the email addresses for each company. While initially unfamiliar with the Salesforce API, I acquired the necessary knowledge to utilize it effectively. By learning about Salesforce Object Query Language (SOQL) queries, I successfully accessed Salesforce data programmatically. To ensure accurate email retrieval, I developed an algorithm that executes multiple SOQL queries to guarantee the correct email address is obtained for each company.

### Email Template Modification
Figuring out how to programmatically modify email templates for automated email sending was another hurdle. By understanding that Outlook email templates are stored in HTML format, I devised a solution to manipulate the HTML content using Python. This involved extracting the HTML content, replacing placeholder text with html elements (an unordered list of drivers), and sending the modified email templates.

These challenges provided valuable learning experiences, allowing me to expand my knowledge and skills in data scraping, API integration, and email automation.

## Overall Impact and Sigificance
The implementation of this project has resulted in significant cost savings and productivity gains for the organization. By automating the email process, the project has not only saved thousands of dollars in labor costs but also freed up valuable employee time. With employees no longer burdened by manual email tasks, they can redirect their focus towards more strategic and value-added activities, ultimately enhancing overall efficiency and productivity within the organization.

## Installation and Usage
This project was developed specifically for Transportation Compliance Service. Thus, attempting to clone the repository and execute the program is highly not recommended.

## Contribution
This project is not open to contribution from others, as it was developed solely to fit the needs of Transportation Compliance Service.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
