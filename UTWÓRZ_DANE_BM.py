import pandas as pd

# Określ plik i kolumny do pobrania
plik = 'wykaz.xlsx'
kolumny = ['BM_REFERENCJA', 'BM_MARSZRUTA', 'BM_TYP']

# Pobierz dane z wybranych kolumn
df = pd.read_excel(plik, usecols=kolumny)

# Zamień nazwy kolumn
df.rename(columns={'BM_REFERENCJA': 'PrdRef', 'BM_MARSZRUTA': 'MARSZRUTA', 'BM_TYP': 'typ'}, inplace=True)

# Dodaj puste kolumny
df['WrkRef1'] = ''
df['OprRef1'] = ''
df['WrkRef2'] = ''
df['OprRef2'] = ''
df['WrkRef3'] = ''
df['OprRef3'] = ''
df['WrkRef4'] = ''
df['OprRef4'] = ''
df['WrkRef5'] = ''
df['OprRef5'] = ''
df['WrkRef6'] = ''
df['OprRef6'] = ''
df['WrkRef7'] = ''
df['OprRef7'] = ''

# Zapisz dane do nowego pliku
df.to_excel('dane.xlsx', index=False)
