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
print()
print("-" * 78)
print("| {:17.17} | {:38.38} | {:4.4} | {:6} |".format("PART NUMBER", "DESCRIPTION", "UOM", "18 QTY"))
print("-" * 78)

for item in options["1200A"]["OUTFITTING PARTS"]:
    if item["PART NUMBER"]:
        print("| {:17.17} | {:38.38} | {:4.4} | {:6.2f} |".format(item["PART NUMBER"], item["DESCRIPTION"], item["UOM"], item["18 QTY"]))

print("-" * 78)
