#!/usr/bin/env python

import pickle

# load data from piclk
with open(r"K:\Links\2020\Options\options.pickle", "rb") as file:
    options =  pickle.load(file)

"""
output some part information to the screen
  for more on formatting see https://pyformat.info/
  for text it is {: how wiide should the column be, pad with spaces as necessary PERIOD at what point to cut off letters if there are to many
  for numbers with decimal points total lenth, including padding and whats after the decimal point PERIOD how many numbers after the decimal point
"""

def print_parts(dept, option):
	title = option + ":" + dept
	width = 87
	center = int(((width - 2) - len(title)) / 2)
	if dept == "FABRICATION":
		print()
		print("-" * width)
	if dept == "CANVAS":
		print("-" * width)
		
	print("|{:85}|".format(" " * center + title))
	print("-" * width)
	print("| {:17.17} | {:38.38} | {:4.4} | {:6} | {:6} |".format("PART NUMBER", "DESCRIPTION", "UOM", "18 QTY", "20 QTY"))
	print("-" * width)
	for item in options[option][dept + " PARTS"]:
		if item["PART NUMBER"]: 
		
		    qty18 = item["18 QTY"]
			if qty18 in ["INPUT", "NA"]:
				qty18 = "0.0"
				
				qty20 = item["20 QTY"]
			if qty20 in ["INPUT", "NA"]:
				qty20 = "0.0"
				
			print("| {:17.17} | {:38.38} | {:4.4} | {:6.2f} | {:6.2f} |".format(item["PART NUMBER"], item["DESCRIPTION"], item["UOM"], qty18, float(item["18 QTY"])))

	if dept == "OUTFITTING":
		print("-" * width)

def print_departments(option):
	for dept in ["FABRICATION", "CANVAS", "PAINT", "OUTFITTING"]:
		print_parts(dept, option)

for option in sorted(options):
# for option in ["1200A", "1200B", "70"]:
	if option in []:
		continue
	print_departments(option)
	print("\n")
	