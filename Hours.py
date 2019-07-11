#!/usr/bin/env python
import pickle

# load data from piclke
with open(r"K:\Links\2020\Options\options.pickle", "rb") as file:
    options =  pickle.load(file)
	
# file = open(r"output\18ft hours.csv", "w")
with open(r"output\book export\18hours.csv", "w") as file:
	for option in options:
			LABOR_HOURS = options[option]["18 LABOR TOTAL"]
			if LABOR_HOURS is None:
				LABOR_HOURS = "0"
			DESIGN_HOURS = options[option]["18 DESIGN HOURS"]
			if DESIGN_HOURS is None:
				DESIGN_HOURS = "0"
			FABRICATION_HOURS = options[option]["FABRICATION 18 HOURS"]
			if FABRICATION_HOURS is None:
				FABRICATION_HOURS = "0"
			CANVAS_HOURS = options[option]["CANVAS 18 HOURS"]
			if CANVAS_HOURS is None:
				CANVAS_HOURS = "0"
			PAINT_HOURS = options[option]["PAINT 18 HOURS"]
			if PAINT_HOURS is None:
				PAINT_HOURS = "0"
			OUTFITTING_HOURS = options[option]["OUTFITTING 18 HOURS"]
			if OUTFITTING_HOURS is None:
				OUTFITTING_HOURS = "0"
			OPTION_NAME = options[option]["OPTION NAME"]
		
file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\19hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["19 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["19 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 19 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 19 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 19 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 19 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\20hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["20 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["20 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 20 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 20 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 20 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 20 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\21hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["21 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["21 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 21 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 21 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 21 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 21 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\22hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["22 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["22 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 22 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 22 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 22 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 22 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\23hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["23 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["23 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 23 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 23 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 23 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 23 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\24hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["24 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["24 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 24 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 24 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 24 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 24 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\25hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["25 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["25 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 25 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 25 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 25 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 25 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\26hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["26 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["26 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 26 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 26 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 26 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 26 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\27hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["27 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["27 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 27 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 27 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 27 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 27 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\28hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["28 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["28 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 28 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 28 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 28 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 28 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\29hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["29 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["29 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 29 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 29 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 29 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 29 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\30hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["30 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["30 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 30 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 30 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 30 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 30 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\31hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["31 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["31 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 31 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 31 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 31 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 31 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\32hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["32 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["32 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 32 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 32 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 32 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 32 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\33hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["33 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["33 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 33 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 33 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 33 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 33 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\34hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["34 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["34 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 34 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 34 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 34 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 34 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\35hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["35 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["35 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 35 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 35 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 35 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 35 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\36hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["36 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["36 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 36 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 36 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 36 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 36 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))

with open(r"output\book export\37hours.csv", "w") as file:
	for option in options:
		LABOR_HOURS = options[option]["37 LABOR TOTAL"]
		if LABOR_HOURS is None:
			LABOR_HOURS = "0"
		DESIGN_HOURS = options[option]["37 DESIGN HOURS"]
		if DESIGN_HOURS is None:
			DESIGN_HOURS = "0"
		FABRICATION_HOURS = options[option]["FABRICATION 37 HOURS"]
		if FABRICATION_HOURS is None:
			FABRICATION_HOURS = "0"
		CANVAS_HOURS = options[option]["CANVAS 37 HOURS"]
		if CANVAS_HOURS is None:
			CANVAS_HOURS = "0"
		PAINT_HOURS = options[option]["PAINT 37 HOURS"]
		if PAINT_HOURS is None:
			PAINT_HOURS = "0"
		OUTFITTING_HOURS = options[option]["OUTFITTING 37 HOURS"]
		if OUTFITTING_HOURS is None:
			OUTFITTING_HOURS = "0"
		OPTION_NAME = options[option]["OPTION NAME"]
		
		file.write("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(option, OPTION_NAME, LABOR_HOURS, DESIGN_HOURS, FABRICATION_HOURS, CANVAS_HOURS, PAINT_HOURS, OUTFITTING_HOURS, OPTION_NAME))
