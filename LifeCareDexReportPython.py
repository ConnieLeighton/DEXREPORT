import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os

# Load the Excel files
billings_path = os.path.join(os.getcwd(), 'Billings_Patient.xlsx')
appointments_path = os.path.join(os.getcwd(), 'Appointments_Patient.xlsx')
codes_path = os.path.join(os.getcwd(), 'CHSP Codes.xlsx')
dds_clients_path = os.path.join(os.getcwd(), 'DSSClients.xlsx')

billings_data = pd.read_excel(billings_path)
appointments_data = pd.read_excel(appointments_path)
codes_data = pd.read_excel(codes_path)
dds_clients_data = pd.read_excel(dds_clients_path)

# Create a map of CHSP Codes for quick lookup
chsp_code_map = {}
for _, code in codes_data.iterrows():
    dex_dss_category = code['DEX DSS Category'].replace('&', 'and').replace('Ongoing Allied Health andTherapy Services', 'Ongoing Allied Health and Therapy Services')
    chsp_code_map[code['Code']] = {
        'ScheduledService': code['Visit Type in HCM'],
        'Minutes': code['Total Time Reported'] if code['Total Time Reported'] != 'as per report' else 'as per report',
        'DEXDSSCategory': dex_dss_category,
        'ServiceTypeId': code.get('Service Type ID', 'Unknown')
    }

# Create a map of DDS Clients for quick lookup
dds_clients_map = {client['DSSClientID']: client for _, client in dds_clients_data.iterrows()}

# Function to convert Excel date to Python date string
def excel_date_to_js_date(excel_date):
    if isinstance(excel_date, (int, float)):
        return (pd.to_datetime('1899-12-30') + pd.to_timedelta(excel_date, 'D')).strftime('%Y-%m-%d')
    elif isinstance(excel_date, pd.Timestamp):
        return excel_date.strftime('%Y-%m-%d')
    elif isinstance(excel_date, str):
        return excel_date  # Assuming it's already in the correct format
    return 'Unknown'

# Load and parse OrganisationData.xml
organisation_data_path = os.path.join(os.getcwd(), 'OrganisationData.xml')
tree = ET.parse(organisation_data_path)
organisation_data = tree.getroot()

# Create a map for ServiceTypes for quick lookup
service_type_map = {}
for activity in organisation_data.findall('.//OrganisationActivity'):
    for service_type in activity.findall('.//ServiceType'):
        service_name = service_type.find('Name').text
        service_type_id = service_type.find('ServiceTypeId').text
        service_type_map[service_name] = service_type_id

# Create the desired object structure
result_data = {
    'Clients': [],
    'Cases': [],
    'Sessions': []
}

case_map = {}
client_id_set = set()

def update_value(value):
    return 'true' if value == 1 else 'false'

# Function to clean up IDs by removing trailing .0 if present
def floatToString(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)

# Add Clients data to the result object first
valid_billings_client_ids = set(billings_data['Client ID'])

for client_id, client in dds_clients_map.items():
    birth_date = excel_date_to_js_date(client['DateOfBirth'])
    
    if client['PracSuiteID'] and client['DSSClientID'] and client['PracSuiteID'] in valid_billings_client_ids:
        if client_id not in client_id_set:
            has_disabilities = client['Disabilities'] != '<NONE>'
            client_node = {
                'ClientId': floatToString(client['PracSuiteID']) or 'no value',
                'Slk': client['SLK'],
                'ConsentToProvideDetails': update_value(client['ConsentToProvideDetails']),
                'ConsentedForFutureContacts': update_value(client['ConsentedForFutureContacts']),
                'GivenName': client['FirstName'] or 'no value',
                'FamilyName': client['LastName'] or 'no value',
                'IsUsingPsuedonym': update_value(client['IsUsingPseudonym']),
                'BirthDate': birth_date,  # Formatted birth date
                'IsBirthDateAnEstimate': update_value(client['IsBirthDateAnEstimate']),
                'GenderCode': client['GenderCode'] or 'NOTSTATED',
                'CountryOfBirthCode':floatToString(client['CountryOfBirthCode']) or '0',
                'LanguageSpokenAtHomeCode': floatToString(client['LanguageSpokenAtHomeCode']) or '2',
                'AboriginalOrTorresStraitIslanderOriginCode': client['AboriginalOrTorresCode'] or 'NOTSTATED',
                'HasDisabilities': has_disabilities
            }

            if has_disabilities:
                client_node['Disabilities'] = {
                    'DisabilityCode': client['Disabilities']
                }

            client_node['AccommodationTypeCode'] = client.get('AccommodationTypeCode', 'NOTSTATED')
            client_node['DVACardStatusCode'] = client.get('DVACardStatusCode', '')
            client_node['HasCarer'] = update_value(client['HasCarer'])
            client_node['ResidentialAddress'] = {
                'AddressLine1': client.get('Address', 'no value'),
                'Suburb': client.get('Town', 'no value'),
                'StateCode': 'SA' if client.get('County') == 'South Australia' else (client.get('County', 'no value')),
                'Postcode': client.get('PostCode', 'no value')
            }
            client_node['HouseholdCompositionCode'] = client.get('HouseholdCompositionCode', 'no value')

            if client['ConsentToProvideDetails'] == 0:
                client_node['Slk'] = client['SLK']
                del client_node['FamilyName']
                del client_node['GivenName']
            else:
                del client_node['Slk']

            result_data['Clients'].append(client_node)
            client_id_set.add(client_id)

unique_time_id = str(int(pd.Timestamp.now().timestamp()))

for _, billing in billings_data.iterrows():
    client_id = billing['Client ID']
    invoice_id = billing['Invoice #']
    service_code = billing['Item']
    schedule = billing['Schedule']
    item_date = billing['Item Date']
    case_id = f"{client_id}_10714"
    fee_category = billing['Fee Category']
    fee = billing['Fee']

    if fee_category != 'CHSP - Payneham':
        continue
    
    if case_id not in case_map:
        new_case = {
            'CaseId': case_id,
            'OutletActivityId': '10714',
            'TotalNumberOfUnidentifiedClients': 0,
            'CaseClients': {
                'CaseClient': {
                    'ClientId': client_id
                }
            }
        }
        result_data['Cases'].append(new_case)
        case_map[case_id] = True

    invoice_data = {}

    service_type_name = chsp_code_map[service_code]['DEXDSSCategory'] if service_code in chsp_code_map else None
    service_type_id = service_type_map.get(service_type_name)

    if schedule == 'CHSP':
        minutes = chsp_code_map[service_code]['Minutes'] if service_code in chsp_code_map else 0
        if minutes == 'as per report':
            appointment = appointments_data[(appointments_data['Client ID'] == client_id) & 
                                            (appointments_data['Appointment Date'] == item_date) & 
                                            (appointments_data['Appointment Status'] != 'Cancelled')]

            minutes = int(appointment['Duration'].values[0]) if not appointment.empty else 0

        invoice_data = {
            'InvoiceID': invoice_id,
            'ScheduledService': chsp_code_map[service_code]['ScheduledService'] if service_code in chsp_code_map else None,
            'FeeCategory': billing['Fee Category'],
            'Minutes': minutes,
            'ServiceTypeId': service_type_id
        }

    elif schedule == 'Occupational Therapy':
        appointment = appointments_data[(appointments_data['Client ID'] == client_id) & 
                                        (appointments_data['Appointment Date'] == item_date) & 
                                        (appointments_data['Appointment Status'] != 'Cancelled')]

        if not appointment.empty:
            invoice_data = {
                'InvoiceID': invoice_id,
                'ScheduledService': appointment['Appointment Type'].values[0],
                'FeeCategory': billing['Fee Category'],
                'Minutes': int(appointment['Duration'].values[0]),
                'ServiceTypeId': service_type_id
            }
        else:
            invoice_data = {
                'InvoiceID': invoice_id,
                'FeeCategory': billing['Fee Category'],
                'Minutes': 0,
                'ServiceTypeId': service_type_id
            }

    if invoice_data.get('InvoiceID'):
        session = {
            'SessionId': unique_time_id,
            'CaseId': case_id,
            'SessionDate': excel_date_to_js_date(item_date),
            'ServiceTypeId': service_type_id,
            'FeesCharged': floatToString(fee),
            'SessionClients': {
                'SessionClient': {
                    'ClientId': client_id,
                    'ParticipationCode': 'CLIENT'
                }
            },
            'TimeMinutes': invoice_data['Minutes']
        }
        result_data['Sessions'].append(session)

    unique_time_id = str(int(unique_time_id) + 1)

# Convert the result object to XML
root = ET.Element('DEXFileUpload')
clients_element = ET.SubElement(root, 'Clients')
cases_element = ET.SubElement(root, 'Cases')
sessions_element = ET.SubElement(root, 'Sessions')

for client in result_data['Clients']:
    client_element = ET.SubElement(clients_element, 'Client')
    for key, value in client.items():
        child_element = ET.SubElement(client_element, key)
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                sub_element = ET.SubElement(child_element, sub_key)
                sub_element.text = str(sub_value)
        else:
            child_element.text = str(value)

for case in result_data['Cases']:
    case_element = ET.SubElement(cases_element, 'Case')
    for key, value in case.items():
        if key == 'CaseClients':
            case_clients_element = ET.SubElement(case_element, 'CaseClients')
            case_client_element = ET.SubElement(case_clients_element, 'CaseClient')
            for client_key, client_value in value['CaseClient'].items():
                client_element = ET.SubElement(case_client_element, client_key)
                client_element.text = str(client_value)
        else:
            case_child_element = ET.SubElement(case_element, key)
            case_child_element.text = str(value)

for session in result_data['Sessions']:
    session_element = ET.SubElement(sessions_element, 'Session')
    for key, value in session.items():
        if key == 'SessionClients':
            session_clients_element = ET.SubElement(session_element, 'SessionClients')
            session_client_element = ET.SubElement(session_clients_element, 'SessionClient')
            for client_key, client_value in value['SessionClient'].items():
                client_element = ET.SubElement(session_client_element, client_key)
                client_element.text = str(client_value)
        else:
            session_child_element = ET.SubElement(session_element, key)
            session_child_element.text = str(value)

# Convert the XML tree to a string
xml_str = ET.tostring(root, 'utf-8')
parsed_str = minidom.parseString(xml_str)
pretty_xml_str = parsed_str.toprettyxml(indent="  ")

# Write the XML to a file
output_path = os.path.join(os.getcwd(), 'result.xml')
with open(output_path, 'w') as f:
    f.write(pretty_xml_str)

print('XML file has been created.')
