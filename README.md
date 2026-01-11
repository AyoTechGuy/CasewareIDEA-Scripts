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

When importing into CasewareIDEA, import should be done line by line and final import should look like below. I always recommend importing all parameters as character

![Image Alt](https://github.com/AyoTechGuy/CasewareIDEA-Scripts/blob/c5c1a54c1eb357d81ea78d1b1af4f8557efd5d7d/Screenshot%202026-01-10%20085114.png)

Once this is done, you can then run the script to merge the descriptions to a single line for each transactions and your final results should look like this

![Image Alt](https://github.com/AyoTechGuy/CasewareIDEA-Scripts/blob/082ae819be732aa2bb2be82e619c8df676ac5ed1/Screenshot%202026-01-11%20070457.png)
