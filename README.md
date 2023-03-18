# Truck Load Invoice Excel Worksheet
## About
I made this script to help simplfy entering truck load ticket data for invoices while working at Adem Inc. (my uncle's trucking company) in 2019.
This Python script creates an Excel worksheet to keep track of truck load invoices. 
The user is prompted to input data for each invoice, including dates, ticket numbers, truck numbers, customer or job names, loads or hours, and rate per load or hour. 
The script calculates the total amount owed for each row and the grand total for all the invoices.

The Excel worksheet has the following columns:

- Dates
- Ticket Numbers
- Truck Number
- Customer/Job
- Loads/Hours
- Rate per Load
- Total Amount Owed

The user is prompted to enter data for each row, and can add as many rows as needed. 
Once all the data has been entered, the script saves the worksheet as an Excel file with the name and directory specified by the user.

## Getting Started

To use this script, you need to have Python 3 installed on your computer. You also need to have the `openpyxl` module installed. You can install it using pip:
```bash
pip install opnepyxl
```



## Usage
To run the script, simply execute it using Python:
```python
python truck_load_invoice.py
```
The script will prompt you for the invoice data and then save the Excel file to the specified directory, along with the name you specify.
