import pprint, openpyxl


print('Opening workbook...')
wb = openpyxl.load_workbook('censuspopdata.xlsx')
sheet = wb['Population by Census Tract']
countryData = {}

print('Reading rows...')
for row in range(2, sheet.max_row + 1):
    # read each row
    state = sheet['B'+str(row)].value
    country = sheet['C'+str(row)].value
    pop = sheet['D'+str(row)].value

    # DataStructure
    # countryData[State abbrev][country]['tracts']
    # countryData[State abbrev][country]['pop']
    countryData.setdefault(state, {})
    countryData[state].setdefault(country, {'tracts': 0, 'pop': 0})
    countryData[state][country]['tracts'] += 1
    countryData[state][country]['pop'] += int(pop)

# write data
print('Wrting results...')
resultFile = open('census2010.py', 'w')
resultFile.write('allData = ' + pprint.pformat(countryData))
resultFile.close()
print('Done.')






