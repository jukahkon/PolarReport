import os
from pprint import pprint

def enumReports(dir):
    reports = []

    for path, subdirs, files in os.walk(dir):
        for name in files:
            print name
            if '.xlsx' in name:
                reports.append(os.path.join(path, name))

    return reports

# TEST
if __name__ == "__main__":
    dir = 'c:\Projects\PolarReport\PolarReport'
    
    reports = enumReports(dir)

    pprint(reports)



