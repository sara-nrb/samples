#!/usr/bin/env python

import pickle

# load data from piclke
with open(r"K:\Links\2020\yamaha rigging\yamaha rigging.pickle", "rb") as file:
    options =  pickle.load(file)
	
# file = open(r"output\eos.csv", "w")

with open(r"output\book export\yamaha rigging.csv", "w") as file:
	for option in options:
		OPTION_NAME = options[option]["OPTION NAME"]
		OPTION_NOTES = options[option]["OPTION NOTES"]
		if OPTION_NOTES is None:
			OPTION_NOTES = ""
		CALCULATED_RETAIL = options[option]["CALCULATED RETAIL"]
		ADVERTISED_RETAIL = options[option]["ADVERTISED RETAIL"]
		if ADVERTISED_RETAIL is None or ADVERTISED_RETAIL == "N/C":
			ADVERTISED_RETAIL = "N/C"
		else:
			ADVERTISED_RETAIL = "{:.2f}".format(float(ADVERTISED_RETAIL))
		
		file.write("{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, CALCULATED_RETAIL, ADVERTISED_RETAIL, OPTION_NOTES))

file.close()


