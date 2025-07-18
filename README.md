# Estimation Script For UNITED GLASS VENTURES 

# To compile to EXE:
#   Enter venv: venv\Scripts\activate
#   run: pyinstaller --onefile --noconsole (only if UI) NAMEOFYOURFILE.py
#   Compiled program will be found in the dist folder
#   exit

# New ui designing needs recompiling and rerelease aka take binary ui and drop down on github

# with data adjust the AS to whatever the user inputs and the fetch withever the user inputs as well

# 1. Main.Py runs and Code Goes Thru GUI player inputs stuff and it gets stored on a variable
# 2. Main.Py calculates system specific outputsbased on inpurts and returns back a calculated_output with the format [{des,quan,part},....]
# 3. Main.Py autocalculates other outputs(Area+Perimeter) and sends all input+outputs(the list of dics) to generate excel report