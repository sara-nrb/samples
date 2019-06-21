#!/usr/bin/env python

import pickle

# load data from piclk
with open(r"K:\Links\2020\Options\options.pickle", "rb") as file:
    options =  pickle.load(file)

# wite output to csv file 
with open(r"output\eos.csv", "w") as file:
    file.write("{}, {}\n".format("OPTION", "DEPARTMENT"))
    for option in sorted(options):
        value = options[option]["EOS DEPARTMENT"]
        if value is None:
            value = ""
        file.write("{}, {}\n".format(option, value))