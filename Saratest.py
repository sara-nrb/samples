#!/usr/bin/env python

import pickle

# load data from piclk

"""
output some part information to the screen
  for more on formatting see https://pyformat.info/
  for text it is {: how wiide should the column be, pad with spaces as necessary PERIOD at what point to cut off letters if there are to many
  for numbers with decimal points total lenth, including padding and whats after the decimal point PERIOD how many numbers after the decimal point
"""

def print_outfitting(part):
	width = 87
	center = int(((width - 2) - len(part)) / 2)
	print()
	print("-" * width)
	print("|{:85}|".format(" " * center + part))
	print("-" * width)
	print("| {:17.17} | {:38.38} | {:4.4} | {:6} | {:6} |".format("PART NUMBER", "DESCRIPTION", "UOM", "18 QTY", "20 QTY"))
	print("-" * width)
	for item in options[part]["OUTFITTING PARTS"]:
		if item["PART NUMBER"]: 
			print("| {:3.3}| {:17.17} | {:38.38} | {:4.4} | {:6.2f} | {:6.2f} |".format("Bob", item["PART NUMBER"], item["DESCRIPTION"], item["UOM"], item["18 QTY"], item["18 QTY"]))

	print("-" * width)

for part in ["1200A", "1200B", "70"]:
	print_outfitting(part)