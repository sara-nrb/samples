#!/usr/bin/env python

import pickle
from openpyxl import load_workbook

# load data from pickle
with open(r"K:\Links\2020\Options\options.pickle", "rb") as file:
	options =  pickle.load(file)

	with open(r"C:\Development\samples\output\Parts Costing for Options 18.xlsx", "w") as file:
		for item in options["OUTFITTING PARTS"]:
			file.write(options[parts]["OUTFITTING PART NUMBER"], options[parts]["OUTFITTING UOM"])
			