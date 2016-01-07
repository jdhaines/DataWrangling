# -*- coding: utf-8 -*-
# Find the time and value of max load for each of the regions
# COAST, EAST, FAR_WEST, NORTH, NORTH_C, SOUTHERN, SOUTH_C, WEST
# and write the result out in a csv file, using pipe character | as the
# delimiter.  An example output can be seen in the "example.csv" file.

import xlrd
import os
import csv
from zipfile import ZipFile

datafile = "data/2013_ERCOT_Hourly_Load_Data.xls"
outfile = "2013_Max_Loads.csv"


def open_zip(datafile):
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()


def parse_file(datafile):
    '''Here is some docstring text.'''
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    # YOUR CODE HERE
    # Remember that you can use xlrd.xldate_as_tuple(sometime, 0) to convert
    # Excel date to Python tuple of (year, month, day, hour, minute, second)
    return sheet


def save_file(sheet, filename):
    '''Get a sheet for xlrd use, data list of lists, and an output filename.'''

    header = ['Station', 'Year', 'Month', 'Day', 'Hour', 'Max Load']

    data = [[sheet.cell_value(r, col) for col in range(sheet.ncols)]
            for r in range(sheet.nrows)]

    # Build the station list
    stations = []
    for i in range(sheet.ncols):
        stations.append(sheet.col_values(i)[0])
    stations.pop()
    stations.reverse()
    stations.pop()
    stations.reverse()

    # Get the column numbers for each station in a dict.
    # {u'FAR_WEST': 3, u'NORTH': 4, u'ERCOT': 9, ...
    col_nums = {}
    for station in stations:
        col_nums.setdefault(station)  # set keys

    for i, station in enumerate(stations):
        if sheet.col_values(i+1)[0] in stations:
            col_nums[station] = i + 1

    # Find the max values for each column.
    maxvals = []
    for i, station in enumerate(stations):
        maxvals.append(max(sheet.col_values(col_nums[station], start_rowx=1)))

    # Find the times for the max values
    maxval_times = []
    for i, station in enumerate(stations):
        for j in range(len(sheet.col_values(i+1))):  # needs equal number rows
            if data[j][i+1] == maxvals[i]:
                maxval_times.append((xlrd.xldate_as_tuple(data[j][0], 0))[:4])

    # Round the max values
    rounded_maxvals = []
    for k in range(len(maxvals)):
        rounded_maxvals.append(round(maxvals[k], 1))

    # Format the output
    output = []
    for i in range(len(stations)):
        row = []
        row.append(stations[i])
        row.extend(list(maxval_times[i]))
        row.append(rounded_maxvals[i])
        output.append(row)
    output.insert(0, header)

    with open(filename, 'wb') as csvfile:
        mywriter = csv.writer(csvfile, delimiter='|')
        for n in range(len(output)):
            mywriter.writerow(output[n])

# data, sheet = parse_file(datafile)
# save_file(data, sheet, outfile)

def test():
    # open_zip(datafile)
    sheet = parse_file(datafile)
    save_file(sheet, outfile)

    number_of_rows = 0
    stations = []

    ans = {'FAR_WEST': {'Max Load': '2281.2722140000024',
                        'Year': '2013',
                        'Month': '6',
                        'Day': '26',
                        'Hour': '17'}}
    correct_stations = ['COAST', 'EAST', 'FAR_WEST', 'NORTH',
                        'NORTH_C', 'SOUTHERN', 'SOUTH_C', 'WEST']
    fields = ['Year', 'Month', 'Day', 'Hour', 'Max Load']

    with open(outfile) as of:
        csvfile = csv.DictReader(of, delimiter="|")
        for line in csvfile:
            station = line['Station']
            if station == 'FAR_WEST':
                for field in fields:
                    # Check if 'Max Load' is within .1 of answer
                    if field == 'Max Load':
                        max_answer = round(float(ans[station][field]), 1)
                        max_line = round(float(line[field]), 1)
                        assert max_answer == max_line

                    # Otherwise check for equality
                    else:
                        assert ans[station][field] == line[field]

            number_of_rows += 1
            stations.append(station)

        # Output should be 8 lines not including header
        assert number_of_rows == 8

        # Check Station Names
        assert set(stations) == set(correct_stations)


if __name__ == "__main__":
    test()