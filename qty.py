#!/usr/bin/env python

import pickle

# load data from piclke
with open(r"K:\Links\2020\Options\options.pickle", "rb") as file:
	options =  pickle.load(file)
	
# file = open(r"output\eos.csv", "w")
with open(r"output\lsob.csv", "w") as file:
	for option in options:
		LSOB = options[option]["LSOB"]
		EOS_DEPARTMENT = options[option]["EOS DEPARTMENT"]
		if EOS_DEPARTMENT is None:
			EOS_DEPARTMENT = ""
		else:
			EOS_DEPARTMENT = EOS_DEPARTMENT[11:]
		OPTION_NAME = options[option]["OPTION NAME"]
		OPTION_NOTES = options[option]["OPTION NOTES"]
		if OPTION_NOTES is None:
			OPTION_NOTES = ""
		ADVERTISED_RETAIL = options[option]["ADVERTISED RETAIL"]
		if ADVERTISED_RETAIL is None or ADVERTISED_RETAIL == "N/C":
			ADVERTISED_RETAIL = "N/C"
		else:
			ADVERTISED_RETAIL = "{:.2f}".format(float(ADVERTISED_RETAIL))
		
		if LSOB == "Y":
			file.write("{}\t{}\t{}\t{}\t{}\n".format(option, sorted(EOS_DEPARTMENT), OPTION_NAME, OPTION_NOTES, ADVERTISED_RETAIL))


