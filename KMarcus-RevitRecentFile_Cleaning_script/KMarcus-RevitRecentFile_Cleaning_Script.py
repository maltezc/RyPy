import os
import re
import getpass
import ConfigParser

username = getpass.getuser()
p = ConfigParser.SafeConfigParser()

os.chdir(r'C:\Users\cmaltez\Desktop\Maltez\PyDy Scripts\RyPy\KMarcus-RevitRecentFile_Cleaning_script')
# os.chdir("/Users/chrismaltez/Desktop/kmarcus")
filename = "Revit.ini"

# filename = "revit.text"
open(filename, 'r')
# with open(filename, 'r') as f:
p.readfp(filename)


# textfile = open(filename, 'r')
xx = "file01 this here as well"
matches = []
pattern = re.findall(r"^\w+", textfile)

for line in textfile:

	print(line)




textfile.close()

