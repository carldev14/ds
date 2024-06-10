from pymongo import MongoClient
from dotenv import load_dotenv
import os
load_dotenv()

mongodb_url = os.getenv("MONGODB_URL")
client = MongoClient(mongodb_url)

mongodb_DB = os.getenv("DATABASE")
db = client[mongodb_DB]

companyCol = os.getenv("CompanyCol")
addressCol = os.getenv("AddressCol")
descriptionCol = os.getenv("DescriptionCol")

company_col = db[companyCol]
address_col = db[addressCol]
description_col = db[descriptionCol]


        
            

def choose_action():
    print('\nAdd Data to MongoDB\n')
    print('[1] Add Company')
    print('[2] Add Address')
    print('[3] Add Description')
    choice = input('\nEnter your choice 1-3: ')
    if choice == '1':
        add_company()
    elif choice == '2':
        add_address()
    elif choice == '3':
        add_description()
    else:
        print('\nInvalid option\n')
        
def add_company():
    while True:
        
        company_add = input('Enter the name of the company: ')
        add = {"name": company_add}
        company_col.insert_one(add)
        if company_add.lower() == 'quit':
            choose_action()
def add_address():
    while True:
        address_add = input('Enter the name of the company: ')
        add = {"name": address_add}
        address_col.insert_one(add)
        if add_address().lower() == 'quit':
            choose_action()    

def add_description():
    while True:
        description_add = input('Enter the name of the company: ')
        add = {"name": description_add}
        description_col.insert_one(add)        
        if add_description().lower() == 'quit':
            choose_action()  
if __name__ == "__main__":

        choose_action()

