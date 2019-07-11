#!/usr/bin/env python

import pickle

# load data from piclk
with open(r"K:\Links\2020\Options\options.pickle", "rb") as file:
    options =  pickle.load(file)

# wite output to csv file 
with open(r"output\parts.csv", "w") as file:
	for option in options:
		for item in options[option]["OUTFITTING PARTS"]:
			if item["PART NUMBER"]: 
				file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(
				option , item["PART NUMBER"], item["DESCRIPTION"], item["UOM"], item["18 QTY"], item["20 QTY"], item["25 QTY"], item["27 QTY"]))