import pandas as pd

base = '/mnt/c/Users/jortega/code/partner-ideal-persona-logic/raw_data'
leads = pd.read_csv(f'{base}/partner type leads.csv')
contacts = pd.read_csv(f'{base}/partner type contacts.csv')

leads.rename(columns={'Company / Account':'Account Name'}, inplace=True)
leads.rename(columns={'Lead ID':'Related Record ID'}, inplace=True)
contacts.rename(columns={'Type': 'Account Type'}, inplace=True)
contacts.rename(columns={'Contact ID': 'Related Record ID'}, inplace=True)

leads['lead/contact'] = 'Lead'
contacts['lead/contact'] = 'Contact'

data = pd.concat([leads, contacts], ignore_index=True)
data['Full Name'] = (
    data['First Name'].fillna('').str.strip() + ' ' +
    data['Last Name'].fillna('').str.strip()
).str.strip()

data.drop(columns=["First Name", "Last Name"], inplace=True)

# data = data[["Full Name", "Title", "Job Level","Account Name","Target Account",\
#              "Account Type","Automated Persona Match","Ideal Persona","Account Owner","Lead Owner","lead/contact","Email","Related Record ID", 'Create Date']]

data = data[['Title', 'Job Level', 'Automated Persona Match', 'Account Type', 'lead/contact', 'Create Date', 'Related Record ID']]

# Quick sanity checks
print(data['lead/contact'].value_counts())
print(data.shape)

data.to_csv('partner_type_people.csv', index=None)
