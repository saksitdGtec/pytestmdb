import pyodbc
msa_driver = [x for x in pyodbc.drivers() if 'ACCESS' in x.upper()]
print(f'MS-Access Drivers : {msa_driver}')