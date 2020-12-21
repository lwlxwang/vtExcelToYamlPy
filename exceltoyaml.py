import os
import pandas as pd
import sys
from pathlib import Path
import yaml
import numpy as np



def filterCSV(filename):
    if filename[-4:] == ".xls":
        return True
    else:
        return False

def strip(value):
    return value.astype(str).str.strip()


class flowmap(dict): pass
def flowmap_rep(dumper, data):
    return dumper.represent_mapping( u'tag:yaml.org,2002:map', data, flow_style=True)

def processfile(filename):
    excelfile = pd.ExcelFile(filename)
    print(excelfile.sheet_names)
    all_data = dict()
    for sheet_name in excelfile.sheet_names:
        df = excelfile.parse(sheet_name)
        # remove first row 
        # use english row for header
        df = pd.DataFrame(
            df.drop(index=0),
            columns=df.iloc[0]
        )
        # remove trailing space
        df = df.apply(strip)
        df = df.replace(np.nan, "", regex=True)
        df = df.replace("nan", "", regex=True)
        # turn into list of dict
        list_of_dict = df.to_dict('records')

        #append to all data
        all_data[sheet_name] = list_of_dict


    yaml.add_representer(flowmap, flowmap_rep)

    for key in all_data:
        all_data[key] = [flowmap(x) for x in all_data[key]]
    
    # print(yaml.dump(all_data))
    result_file = filename[:-4]+".yaml";
    os.makedirs("output", exist_ok=True)
    result_file = os.path.join("./output/", result_file)
    with open(result_file, "w") as writer:
        writer.write(yaml.dump(all_data,width=1000))



filename = ""

path = os.path.dirname(os.path.realpath(__file__))
files = os.listdir(path)
filtered_files = filter(filterCSV, files)

for csvfile in filtered_files:
    print(csvfile)
    processfile(csvfile)



