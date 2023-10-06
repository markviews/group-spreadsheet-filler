# group-spreadsheet-filler
Randomly assigns groups to everyone that doesn't manually join one.<br>
Useful to form groups when some people know who they want to work with but others want to be randomly assigned.

#### How to use
1. create google spreadsheet in [this format](https://docs.google.com/spreadsheets/d/1SYSyAxJ16ZBlbt-OUwovlP7WiWaWnHh86ddpXZSyiBw/edit?usp=sharing). You can type any text in the top row and left collum the program does not read them.
3. give everyone edit access to the spreadsheet, ask them to fill in their email or something else unique that you agreed on
4. run the script
   
#### Running the script
1. install [Python](https://www.python.org/downloads/) and libraries `pip install openpyxl`
1. download `fill.py`, edit the config options starting at line #4
2. download your spreadsheet to the same folder
3. create `names.txt` in the same folder with every person's email or unique string on a new line
5. run `python fill.py`
6. If any names appear in the spreadsheet that aren't in your names.txt you will be asked to review them and run the script again
