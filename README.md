# CasewareIDEA-Scripts
This IDEAScript automates the process of merging multi-line bank statement (especially Barclays bank statement) entries into single, 
consolidated transaction records. It exports data from IDEA to Excel, 
dynamically injects and runs a VBA macro to group related lines based on transaction types 
as trigger words for a new line (e.g., debit, counter credit, standing order), preserves the original 
text format of date fields, and reimports the merged results back into IDEA as a new database.
The script improves efficiency and consistency in cleaning and structuring
bank transaction data for further analysis.
Some Barclays bank statements are laid out like this below
![Image Alt](https://github.com/AyoTechGuy/CasewareIDEA-Scripts/blob/bb04732ce172671b4ce6c119427e72c3adab304e/Screenshot%202026-01-10%20083653.png)
