# Invoice-Generator
This python script uses openpyxl to create invoices in Excel format based on a given invoice format


Prerequisites:
Install openpyxl using "pip3 install openpyxl"
Install num2words using "pip3 install num2words"


Instructions:
Edit the Invoice Format Excel file to replace the company name, company field and other such fields with your own relevant details.
Run the python script and input the details asked such as client name,client address, invoice number and date (leave empty for today's date).
Enter the number of entries of items to be done in the invoice. For example, if two items are to be included in the bill, input 2. The max limit of a bill is 25 entries. 
Enter the item name, price and quantity
Enter the GST(Tax) percentage. For example, enter 18% if the total tax on the products is 18%. The script will automatically distribute the GST% into CGST and SGST equally.
Your invoice will be saved in your folder with the name "Invoice No xx.xlsx" with xx representing the invoice number you entered previously. 
