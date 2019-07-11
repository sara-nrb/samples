#!/usr/bin/env python

import pickle

# load data from piclk
with open(r"K:\Links\2020\Options\options.pickle", "rb") as file:
    options =  pickle.load(file)

# wite output to csv file 
with open(r"output\parts.csv", "w") as file:
	for option in options:
		file.write("{}\t{}\t{}\t{}\t{}\n".format(
		option , options[option]["OUTFITTING 18 HOURS"], options[option]["OUTFITTING 20 HOURS"], options[option]["OUTFITTING 25 HOURS"], options[option]["OUTFITTING 27 HOURS"]))