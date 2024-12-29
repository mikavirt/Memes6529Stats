# Memes6529Stats
Extract collection and ownership data from the memes by 6529

Step 1 : Request an API key from OpenSea (3min) : https://docs.opensea.io/reference/api-keys

Step 2 : install Python on your machine
- Go to https://www.python.org/downloads/ , download and install on your machine

Step 3 : install pandas
For Windows:
Open Command Prompt: You can do this by searching for "cmd" in the Start menu and opening it.
Run the Install Command: Type the following command and press Enter:

    pip3 install requests pandas openpyxl

This command will install the requests, pandas, and openpyxl libraries, which are required for your script.

For macOS:
Open Terminal: You can find Terminal in your Applications under Utilities or by searching for it using Spotlight.
Run the Install Command: Type the same command into the Terminal:

    pip3 install requests pandas openpyxl

Step 4 : save the 3 .py files from this repo on your local machine

Step 5 : use an IDE (like IDLE) to open the 3 .py files and 
- Add your API key in the global variables in each file
- Add your wallet addresses to the file "Extract_own_collection.py"
-   if you only have 1 or 2 addresses, you can safely remove the other ones. Same if you have more, just add them
- Save the files

Step 6 : Run the scripts
- Through your Terminal, go to the folder where your 3 .py files are saved
- Type+Enter : "python3 Extract_own_collection.py"
- Type+Enter : "python3 Extract_listings_bid.py"
- Type+Enter : "python3 Extract_names.py"

Finally, if no errors, an excel file should have been created in the same folder with the 3 corresponding sheets.

Feel free to add a summary sheet or additional columns for calculations : they should be kept even when refreshing the data.
