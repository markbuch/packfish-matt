from imap_tools import MailBox, AND
from bs4 import BeautifulSoup
from copy import deepcopy
import config as cfg
import csv
import datetime
import logging

# define global variables
results = []  # holds a list of applicants retrieved from the email form.
applicant = {}  # Used to hold applicant data after extracting it from the email form
fields = list(cfg.Header.fields.values()) # List of column headers for csv file
CERTIFYING_VES = 'CERTIFYING_VES'

# Setup logging and logfile
log_filename = datetime.datetime.now().strftime("%m%d%Y_%H%M%S") + "_script_trace.log"
logging.basicConfig(filename=log_filename, level=logging.DEBUG)


# Functions
def set_exams(exams):
    # set_exams receives a list of exams the applicant is interested in taking.  The exams list will be parsed and
    # the correct columns for the elements will be set.
    # If element 3 and/or element 4 are selected, these selections should show up on the Applicant's
    # registration form in Session Manager.

    # Define column heading names to be used with the csv file.
    req_element_3 = 'REQUESTED_ELEMENT_3'
    req_element_4 = 'REQUESTED_ELEMENT_4'

    # loop through each exam and set corresponding element name and value.
    for exam in exams:
        print(f'exam is: {exam}')
        match exam:
            case 'Element 2 (Technician)':
                if len(exams) == 1:
                    applicant[cfg.Header.fields[req_element_3]] = False
                    applicant[cfg.Header.fields[req_element_4]] = False
            case 'Element 3 (General)':
                applicant[cfg.Header.fields[req_element_3]] = True
                applicant[cfg.Header.fields[req_element_4]] = False
            case 'Element 4 (Amateur Extra)':
                if len(exams) == 1:
                    applicant[cfg.Header.fields[req_element_3]] = False
                    applicant[cfg.Header.fields[req_element_4]] = True
                else:
                    applicant[cfg.Header.fields[req_element_4]] = True

    return None


def export_results_to_csv():
    print('Enter export results to csv function')
    print(f'fields: {fields}')
    print(f'Results: {results}')
    csv_filename = datetime.datetime.now().strftime("%m%d%Y_%H%M%S") + "_session_import.csv"
    logging.info(f'Exporting application results to file: {csv_filename}.')
    # open file and write
    with open(csv_filename, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fields)
        writer.writeheader()
        for item in results:
            writer.writerow(item)
    logging.info('Finished exporting results to csv file.')


def add_certifying_ves_to_applicant_data():
    delim = '~'
    ve1 = cfg.VE.one.upper()
    ve2 = cfg.VE.two.upper()
    ve3 = cfg.VE.three.upper()

    applicant[cfg.Header.fields[CERTIFYING_VES]] = ve1 + delim + ve2 + delim + ve3


# main code
def main():
    logging.info(f'Logging into Mail Server: {cfg.Mail.server}.')
    mb = MailBox(cfg.Mail.server).login(cfg.Mail.user, cfg.Mail.password)

    logging.info('Fetching application registration forms from mail server.')
    messages = mb.fetch(criteria=AND(from_="burst@emailmeform.com", seen=False), mark_seen=True, bulk=True)

    # Start processing retrieved applications
    logging.info('Application processing started...')
    logging.info('Processing first applicant')
    for msg in messages:
        # logging.info('Processing next applicant')
        # parse html email message
        bs = BeautifulSoup(msg.html, 'html.parser')
        # emailme form data resides in a html table.
        # get all table rows
        table_rows = bs.find('table').findAll('tr')
        # process each row to get the table data.
        for row in table_rows:
            print(f'row: {row}')
            name = ""
            value = ""
            td = row.findAll('td')

            # parse each row to get the field name and field value.
            # field name will contain '*:' and the will need to be stripped out.
            # the first td item should be the field name with *:, ex first name*:
            # the second td item should be the value
            logging.info('\n')
            for item in td:
                if item.text.find("*:") != -1:
                    name = item.text.strip().replace("*:", "")
                else:
                    value = item.text.strip()
            # before adding the name/value pair to the applicant dict, check for required modifications.
            logging.info(f'Performing pre-checks on: {name}, {value} ')

            match name:
                case 'Middle Initial':
                    if value.upper() == 'NONE':
                        # If Middle Initial is NONE, set value to empty string
                        logging.info('Middle Initial was set to NONE, replacing with empty string')
                        applicant[cfg.Header.fields[name]] = ''
                case 'Suffix':
                    if value.upper() == 'NONE':
                        # If Suffix is NONE, set value to empty string
                        logging.info('Suffix was set to NONE, replacing with empty string')
                        applicant[cfg.Header.fields[name]] = ''
                case 'Street Address':
                    if value.find('PO') == -1:
                        # Not a PO Box, add empty PO Box entry
                        applicant[cfg.Header.fields[name]] = value
                        applicant[cfg.Header.fields['PO Box']] = ''
                    else:
                        # PO Box, set Street Address(value) to empty string and add PO_BOX
                        applicant[cfg.Header.fields['PO Box']] = value
                        applicant[cfg.Header.fields[name]] = ''
                case 'Callsign':
                    if value.upper() == 'NOCALL':
                        # If callsign is NOCALL, set Callsign value to empty string and set UPGRADE_LICENSE to False
                        logging.info('Callsign was sent to NONE, replacing with empty string.')
                        applicant[cfg.Header.fields[name]] = ''
                        applicant[cfg.Header.fields['UPGRADE_LICENSE']] = False
                    elif value.isalnum():
                        # If a callsign was entered, set UPGRADE_LICENSE to True and convert callsign to upper case
                        logging.info(f'Callsign: {value.strip()} was detected, setting UPGRADE_LICENSE to true.')
                        applicant[cfg.Header.fields['UPGRADE_LICENSE']] = True
                        applicant[cfg.Header.fields[name]] = value.upper()
                    else:
                        # callsign was entered wrong, needs checked
                        applicant[cfg.Header.fields[name]] = 'ERROR'
                        applicant[cfg.Header.fields['UPGRADE_LICENSE']] = False
                case 'Exams':
                    # Add exams to applicant data
                    set_exams(value.split(', '))
                    # add exams to Notes field
                    applicant[cfg.Header.fields[name]] = value
                case 'City':
                    # Correct formatting.  Capitalize first letter only.
                    logging.info(f'Converting City: {value.strip()} to capitalize first letter only.')
                    applicant[cfg.Header.fields[name]] = value.lower().capitalize()
                case 'State':
                    # Convert state to all upper case.
                    logging.info(f'Converting State: {value.strip()} to upper case.')
                    applicant[cfg.Header.fields[name]] = value.upper()
                case other:
                    # No changes needed, write current values
                    logging.info(f'{value} : No changes were needed.')
                    applicant[cfg.Header.fields[name]] = value

            logging.info(f'Name: {name}, Value: {value}')
            logging.info('-' * 40)

        # Default PREVIOUS_APPLICATION to No
        applicant[cfg.Header.fields['PREVIOUS_APPLICATION']] = 'No'
        # Add certifying VEs to applicant's data.
        add_certifying_ves_to_applicant_data()
        logging.info('Printing Applicant data...\n')
        logging.info(f'Applicant: {applicant}\n')
        # print('Print applicant data')
        # print(applicant)
        # print('Add application to results list')
        logging.info('Adding applicant to results file.')
        results.append(deepcopy(applicant))
        applicant.clear()
        logging.info('Processing next applicant')

    # Logout of mailbox
    logging.info('Logging out of Mailbox.')
    mb.logout()
    # print('Results List')
    # print(results)
    # print('Export applicants to csv')
    export_results_to_csv()


if __name__ == '__main__':
    logging.info(f'{datetime.datetime.now()} ===================================')
    logging.info('Application export process starting')

    main()
    logging.info('Application export process completed.')
    logging.info(f'{datetime.datetime.now()} ===================================')
