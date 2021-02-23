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

def print_outfitting(option):
	width = 102
	center = int(((width - 2) - len(option)) / 2)
	print()
	print("-" * width)
	print("|{:100}|".format(" " * center + option))
	print("-" * width)
	print("| {:12.12} | {:17.17} | {:38.38} | {:4.4} | {:6} | {:6} |".format("OPTION", "PART NUMBER", "DESCRIPTION", "UOM", "18 QTY", "20 QTY"))
	print("-" * width)
	for item in options[option]["OUTFITTING PARTS"]:
		if item["PART NUMBER"]: 
			print("| {:12.12} | {:17.17} | {:38.38} | {:4.4} | {:6.2f} | {:6.2f} |".format(option , item["PART NUMBER"], item["DESCRIPTION"], item["UOM"], item["18 QTY"], item["18 QTY"]))

	print("-" * width)

for option in sorted(options):
	print_outfitting(option)