#!/usr/bin/env python

import pickle

# load data from piclke
with open(r"K:\Links\2020\Options\options.pickle", "rb") as file:
    options =  pickle.load(file)
	
with open(r"output\book export\lsob.csv", "w") as file:
	for option in options:
		LSOB = options[option]["LSOB"]
		LSOB_CODE = options[option]["LSOB CODE"]
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
			file.write("{}\t{}\t{}\t{}\t{}\t{}\n".format(option, EOS_DEPARTMENT, LSOB_CODE, OPTION_NAME, ADVERTISED_RETAIL, OPTION_NOTES))

file.close()

with open(r"output\book export\ssob.csv", "w") as file:
	for option in options:
		SSOB = options[option]["SSOB"]
		SSOB_CODE = options[option]["SSOB CODE"]
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
		
		if SSOB == "Y":
			file.write("{}\t{}\t{}\t{}\t{}\t{}\n".format(option, EOS_DEPARTMENT, SSOB_CODE, OPTION_NAME, ADVERTISED_RETAIL, OPTION_NOTES))

file.close()

with open(r"output\book export\shht.csv", "w") as file:
	for option in options:
		SHHT = options[option]["SHHT"]
		SHHT_CODE = options[option]["SHHT CODE"]
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
		
		if SHHT == "Y":
			file.write("{}\t{}\t{}\t{}\t{}\t{}\n".format(option, EOS_DEPARTMENT, SHHT_CODE, OPTION_NAME, ADVERTISED_RETAIL, OPTION_NOTES))

file.close()

with open(r"output\book export\23os.csv", "w") as file:
	for option in options:
		OS23 = options[option]["23OS"]
		OS23_CODE = options[option]["23OS CODE"]
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
		
		if OS23 == "Y":
			file.write("{}\t{}\t{}\t{}\t{}\t{}\n".format(option, EOS_DEPARTMENT, OS23_CODE, OPTION_NAME, ADVERTISED_RETAIL, OPTION_NOTES))

file.close()

with open(r"output\book export\so.csv", "w") as file:
	for option in options:
		SO = options[option]["SO"]
		SO_CODE = options[option]["SO CODE"]
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
		
		if SO == "Y":
			file.write("{}\t{}\t{}\t{}\t{}\t{}\n".format(option, EOS_DEPARTMENT, SO_CODE, OPTION_NAME, ADVERTISED_RETAIL, OPTION_NOTES))

file.close()

with open(r"output\book export\wxl.csv", "w") as file:
	for option in options:
		WXL = options[option]["WXL"]
		WXL_CODE = options[option]["WXL CODE"]
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
		
		if WXL == "Y":
			file.write("{}\t{}\t{}\t{}\t{}\t{}\n".format(option, EOS_DEPARTMENT, WXL_CODE, OPTION_NAME, ADVERTISED_RETAIL, OPTION_NOTES))

file.close()

with open(r"output\book export\waso.csv", "w") as file:
	for option in options:
		WASO = options[option]["WASO"]
		WASO_CODE = options[option]["WASO CODE"]
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
		
		if WASO == "Y":
			file.write("{}\t{}\t{}\t{}\t{}\t{}\n".format(option, EOS_DEPARTMENT, WASO_CODE, OPTION_NAME, ADVERTISED_RETAIL, OPTION_NOTES))

file.close()

with open(r"output\book export\dv.csv", "w") as file:
	for option in options:
		DV = options[option]["DV"]
		DV_CODE = options[option]["DV CODE"]
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
		
		if DV == "Y":
			file.write("{}\t{}\t{}\t{}\t{}\t{}\n".format(option, EOS_DEPARTMENT, DV_CODE, OPTION_NAME, ADVERTISED_RETAIL, OPTION_NOTES))

file.close()

with open(r"output\book export\c.csv", "w") as file:
	for option in options:
		C = options[option]["C"]
		C_CODE = options[option]["C CODE"]
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
		
		if C == "Y":
			file.write("{}\t{}\t{}\t{}\t{}\t{}\n".format(option, EOS_DEPARTMENT, C_CODE, OPTION_NAME, ADVERTISED_RETAIL, OPTION_NOTES))

with open(r"output\book export\osp.csv", "w") as file:
	for option in options:
		OSP = options[option]["OSP"]
		OSP_CODE = options[option]["OSP CODE"]
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
		
		if OSP == "Y":
			file.write("{}\t{}\t{}\t{}\t{}\t{}\n".format(option, EOS_DEPARTMENT, OSP_CODE, OPTION_NAME, ADVERTISED_RETAIL, OPTION_NOTES))

file.close()

with open(r"output\book export\s.csv", "w") as file:
	for option in options:
		S = options[option]["S"]
		S_CODE = options[option]["S CODE"]
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
		
		if S == "Y":
			file.write("{}\t{}\t{}\t{}\t{}\t{}\n".format(option, EOS_DEPARTMENT, S_CODE, OPTION_NAME, ADVERTISED_RETAIL, OPTION_NOTES))

file.close()


