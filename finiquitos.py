import os
import pandas as pd
from icecream import ic
from Multiherramienta import *
from exchangelib import Credentials, Account, HTMLBody, DELEGATE

def download_scans(only_unread=True):
    path_download = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\DOCUMENTOS\\FINIQUITOS\\scan\\"

    # Set up credentials
    username = 'fprado@javer.com.mx'
    password = 'pamf900509HFA02'
    server = 'https://outlook.office365.com/EWS/Exchange.asmx'  # Exchange server URL

    credentials = Credentials(username=username, password=password)

    # Connect to the Exchange account
    account = Account(primary_smtp_address=username, credentials=credentials, autodiscover=False, access_type=DELEGATE, config=dict(auth_type='basic', server=server))

    # Get the inbox folder
    inbox = account.inbox
    sca_folder = inbox / 'CAO' / 'SCA'

    # Filter unread messages if necessary
    messages = sca_folder.filter(is_read=not only_unread)

    counter = 1
    data = []

    for message in messages:
        try:
            subject = message.subject
            attachments = message.attachments
            raw_date = message.datetime_received
            date = raw_date.strftime("%Y-%m-%d")
            low_subject = subject.lower()

            # Check only emails with attachments
            if attachments:
                for attachment in attachments:
                    filename = attachment.name.lower()

                    # Choose a convenient name, discarding automatic subjects from the scanner
                    if subject.startswith("Message from"):
                        name = filename
                    else:
                        name = low_subject + os.path.splitext(filename)[1]
                    output_file = os.path.join(path_download, name)

                    # Avoid errors from duplicate names
                    try:
                        with open(output_file, 'wb') as f:
                            f.write(attachment.content)
                    except FileExistsError:
                        print(name, "duplicate")
                        # Avoid other possible duplicate names
                        counter_2 = 1
                        while counter_2 <= 20:
                            file, extension = os.path.splitext(name)
                            name = file + "_" + str(counter_2) + extension
                            print(counter_2, name)
                            try:
                                output_file = os.path.join(path_download, name)
                                with open(output_file, 'wb') as f:
                                    f.write(attachment.content)
                                break
                            except FileExistsError:
                                counter_2 += 1
                    except Exception as e:
                        print(e)

                # Save data for a future DataFrame
                data.append({
                    "Subject": low_subject,
                    "Date": date,
                    "Filename": filename,
                    "Output": name})

                print(counter, "-", name, "-", date)
                counter += 1

            # Mark as unread if the message was marked as unread
            if not only_unread and message.is_read:
                message.is_read = False

        except Exception as e:
            print(e)

    # Create and save the DataFrame to an Excel file
    df = pd.DataFrame(data)
    df.to_excel(os.path.join(path_download, "scans.xlsx"), index=False)

    # Close the connection to the Exchange account
    account.logout()


def download_legal_contracts_new(ruta_df, headless=True, ci=False, cm=True, cf=True):
	driver = create_driver(headless=headless)
	df = pd.read_excel(ruta_df, dtype={'Frentes': object})
	contratos = df['Contrato']
	for contrato in contratos:
		conjunto = df['Conjunto'].loc[df['Contrato'] == contrato].values[0]
		ic(contrato, conjunto)
		get_contract_contracts(driver, conjunto, contrato)
		contract_table = beautiful_table(driver)
		table_lenght = len(contract_table) - 1
		ic(table_lenght)
		for i, contract in contract_table.iterrows():
			link = 'LCDocT:DocUrlLink:' + str(i)
			if i == 0:
				if ci:
					swdw(driver, 1, 1, link).click()
					swdw(driver, 1, 1, link)
			elif i == table_lenght:
				if cf:
					swdw(driver, 1, 1, link).click()
					swdw(driver, 1, 1, link)
			else:
				if cm:
					swdw(driver, 1, 1, link).click()
					swdw(driver, 1, 1, link)
			ic(contract)

		swdw(driver, 1, 1, "Return").click()

	# get_contract(driver, conjunto, contract)


# Call the function
# download_scans(only_unread=False)
ruta_df = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\Descarga_finiquitos.xlsx"
