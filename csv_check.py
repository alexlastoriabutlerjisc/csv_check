import time
from tkinter import filedialog
from tkinter import Tk
#import glob
import os
import pandas as pd
import numpy as np
from tkinter.filedialog import askopenfilename
import zipfile



latestyear = '2021/22'


# State colours for indication of errors etc
class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    WHITE  = '\033[0m'  # white (normal)
    RED  = '\033[31m' # red
    GREEN  = '\033[32m' # green
    ORANGE  = '\033[33m' # orange
    BLUE  = '\033[34m' # blue
    PURPLE  = '\033[35m' # purple
    r = 'r'
    g = 'g'
    b = 'b'
    

def getdir():
    fileslist = []
    root = Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory()
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.csv') or file.endswith('.zip'):
                fileslist.append(os.path.join(root, file))
    return folder_path, fileslist
csvdir, csvzipfiles = getdir()

# Subfunction to automatically find and remove the header if present ==========
def splitmetadata2(df):
    if 'UKPRN' or 'Academic' or 'Number' or 'Percent' not in df.columns:
        # Top 50 rows
        dfhead = df.head(50)
        # Header typically contains blank cells, replace these with marker value
        dfhead = dfhead.replace(r'^\s*$', 'BLANKCELL', regex=True)
        # Find first row without na/BLANKCELL
        header_row = dfhead[~dfhead.apply(
            lambda row: row.str.contains('BLANKCELL', case=False).any(), axis=1)].index[0]
        metadata = df[:header_row]
        dataframe = df[header_row:]
        colnames = dataframe.iloc[0] # First row of data as headers
        dataframe = dataframe[1:] # Cut data after the first row
        dataframe.columns = colnames # Set the header row as the header 
        # First row becomes column names, revert
        metadata = pd.DataFrame(np.vstack([metadata.columns, metadata]))
        return metadata, dataframe

def contentcheck(contentdf, value, matchtype, checktype, htmltype):
    if matchtype == 'exact':
        checkcount = contentdf.eq(value).sum().sum()
    else:
        checkcount = (contentdf.applymap(lambda x: value.upper() in str(x).upper()).values == True).sum()
    if (checkcount>0):
        print(f"{checktype}{str(checkcount)}\033[0m {matchtype} matches for {value}")
        f.write('<%(type)s>%(str)s</%(type)s> matches for %(val)s<br>' % 
                {'type': htmltype, 'str': checkcount, 'val': value})

#for csvfile in csvzipfiles:
def csvfilecheck(csvfile, filename):   
    
    # First letter capital
    if(os.path.basename(filename)[0].isupper()):
        print('\033[33mFirst letter of file is a capital\033[0m')
        f.write('<%(type)s>%(str)sFirst letter of file is a capital</%(type)s><br>' % 
                {'type': 'r', 'str': os.path.basename(filename)[0]})
    
    df = pd.read_csv(csvfile, converters={i: str for i in range(255)}) # bypass reading nan as nonvalue
    total=len(df)
    # Excel limit check
    if total>1048576: 
        print('\033[31m'+str(total)+'\033[0m rows in '+filename.rsplit('/', 1)[-1])
        f.write('<%(type)s>%(str)s</%(type)s> rows in %(val)s<br>' % 
                {'type': 'r', 'str': str(total), 'val': filename.rsplit('/', 1)[-1]})
    else:
        print('\033[32m'+str(total)+'\033[0m rows in '+filename.rsplit('/', 1)[-1])
        f.write('<%(type)s>%(str)s</%(type)s> rows in %(val)s<br>' % 
                {'type': 'g', 'str': str(total), 'val': filename.rsplit('/', 1)[-1]})
    metadata, data = splitmetadata2(df)
    print('Control check on "All"')
    contentcheck(data, 'All', 'exact', bcolors.BLUE, bcolors.b)
    contentcheck(data, latestyear, 'exact', bcolors.BLUE, bcolors.b)
    contentcheck(data, 'check', 'exact', bcolors.RED, bcolors.r)
    contentcheck(data, '#', 'exact', bcolors.RED, bcolors.r)
    contentcheck(data, '#Error', 'exact', bcolors.RED, bcolors.r)
    contentcheck(data, '#N/A', 'exact', bcolors.RED, bcolors.r)
    contentcheck(data, 'NA', 'exact', bcolors.RED, bcolors.r)
    contentcheck(data, 'NaN', 'exact', bcolors.RED, bcolors.r)
    contentcheck(data, 'nan', 'exact', bcolors.RED, bcolors.r)
    contentcheck(data, '#Div', 'exact', bcolors.RED, bcolors.r)
    contentcheck(data, '#DIV/0!', 'exact', bcolors.RED, bcolors.r)
    contentcheck(data, 'null', 'exact', bcolors.RED, bcolors.r)
    contentcheck(data, 'Null', 'exact', bcolors.RED, bcolors.r)
    contentcheck(data, 'nan', 'exact', bcolors.RED, bcolors.r)
    
    f.write("<p>Metadata summary</p>")
    f.write(metadata.iloc[:,[0,1]][metadata.iloc[:, 0].str.len() > 0].to_html(header=False, index=False))
    
    f.write("<p>Check for years in file</p>")
    try:
        data.columns = data.columns.str.lower()
        f.write(data.groupby('academic year').count().to_html())
    except:
        f.write("<p>Academic year column not found in data</p>")
    
    
    
# Set up HTML report ====
style = """<style type='text/css'>
html {
  font-family: Courier;
}
r {
  color: #ff0000;
}
g {
  color: #006400;
}
b {
  color: #0000ff;
}
</style>"""
f = open(f'csv_report{time.strftime("%Y%m%d")}.html', 'w')
f.write('<html>')
f.write(style)
f.write(f'<p>{time.strftime("%Y/%m/%d %H:%M:%S")}</p>')
f.write(f'<p>CSV check for {csvdir}</p>')
f.write(f'<p>{str(len(csvzipfiles))} files found</p>')
    

for csvfile in csvzipfiles:
    print(csvfile.rsplit('/', 1)[-1] , '*'*(80-len(csvfile.rsplit('/', 1)[-1])))
    f.write(f"<p>{csvfile.rsplit('/', 1)[-1].center(100, '*')}</p>")
    
    # Space in file check
    if ' ' in os.path.basename(csvfile): 
        print('\033[31m'+os.path.basename(csvfile)+'\033[0m contains spaces.')
        f.write('<%(type)s>%(str)s</%(type)s> contains spaces<br>' % 
                {'type': 'r', 'str': os.path.basename(csvfile)})
    
    # Check file contents
    if csvfile.endswith('.zip'):
        with zipfile.ZipFile(csvfile, 'r') as zip_file:
            # Loop through each file in the ZIP archive
            for file_name in zip_file.namelist():
                # Check if the file is a CSV (you can use a more specific check if needed)
                if file_name.endswith('.csv'):
                    print(file_name.rsplit('/', 1)[-1] , '*'*(80-len(file_name.rsplit('/', 1)[-1])))
                    f.write(f"<p>{file_name.rsplit('/', 1)[-1].center(100, '*')}</p>")
                    with zip_file.open(file_name) as csv_file:
                        csvfilecheck(csv_file, file_name)
    else:
        csvfilecheck(csvfile, csvfile)
    
    
f.write(f"<p>{' END '.center(100, '*')}</p>")
f.write('</html>')
f.close()
