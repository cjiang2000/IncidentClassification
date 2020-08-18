import nltk
import pandas as pd
import numpy as np
from operator import itemgetter
import easygui

#choose file to analyze
path = easygui.fileopenbox()
##Read Excel to dataframe
#System List
services = pd.read_excel("SystemList.xlsx", usecols = "N,O,P,Q,R,S,T,U,V,W")
services = services.astype(str)
#Data
data = pd.read_excel(path, sheet_name = 'Combined', usecols = "X,C,T,H,R,Y,AE")
#Mapped System
pasa = pd.read_excel(path,sheet_name = 'Combined', usecols = "AP")
#Incident ID, Filter, and Has SRT
incid = pd.read_excel(path,sheet_name = 'Combined', usecols = "E,BN,BO,BS")
#Array and dictionary for each category of data (mef, before, middle, end, and special)
mef = []
dic_mef = {}
bef = []
dic_b = {}
mid = []
dic_m = {}
end = []
dic_e = {}
special = []
dic_s = {}
#Array of mapped systems we ignore
msbm = ["Applications", 0, "Mainframe", "Data / Database", "Server Infrastructure", "Service Desk"]
#For loop that reads all the data from the dataframes into the dictionaries and arrays created earlier
for index, row in services.iterrows():
	if not pd.isnull(row["Before"]) and row["Before"] != "nan":
		bef.append(row["Before"])
		dic_b[row["Before"].upper()] = row["Match_b"]
	if not pd.isnull(row["Middle"]) and row["Middle"] != "nan":
		mid.append(row["Middle"])
		dic_m[row["Middle"].upper()] = row["Match_m"]
	if not pd.isnull(row["End"]) and row["End"] != "nan":
		end.append(row["End"])
		dic_e[row["End"].upper()] = row["Match_e"]
	if not pd.isnull(row["Special"]) and row["Special"] != "nan":
		special.append(row["Special"])
		dic_s[row["Special"]] = row["Match_s"]
	if not pd.isnull(row["mef"]) and row["mef"] != "nan":
		mef.append(row["mef"])
		dic_mef[row["mef"].upper()] = row["Match_mef"]

##Analyzing the data
#List of column names
listo = list(data.columns)
#Dictionary to prioritize columns
sorto = {"Affected CI": 1,  "Escalation Teams": 2, "Assignment Group": 3,"Solution": 4, "Title": 5, "Update Action": 6,  "Description": 7}
#Dictionary to store the output
duplicates = {}
#Dictionary to keep track of correct matches
yesno = {}

#Analyze MEF, goes through each index and tokenizes each column from each index. If a token is found it is saved alongside it's accompanying data in duplicates.
for index, row in data.iterrows():
	for x in range(len(row)):
		if not pd.isnull(row[x]) and isinstance(row[x], str):
			tokens = nltk.word_tokenize(row[x].replace('-', ' '))
			for y in range(len(tokens)):
				g = tokens[y]
				tokens[y] = tokens[y].upper()
				if tokens[y] in (name.upper() for name in mef):
					if index not in duplicates:
						if pasa.at[index,'Mapped System'] in msbm:
							yesno[index] = ['n/a', 'n/a']
						else:
							yesno[index] = ['n', 'n']
						duplicates[index] = [(0,pasa.at[index,'Mapped System'])]
					temp = ""
					for i in range(10):
						if y - 5 + i >= 0 and y - 5 + i < len(tokens):
							temp = temp + " " + tokens[y-5+i]
					duplicates[index].append((sorto[listo[x]], dic_mef[tokens[y]] + " " + "{" + g + "}" + " " + listo[x] + ":: " + temp, dic_mef[tokens[y]]))

#Analyze Bef
for index, row in data.iterrows():
	for x in range(len(row)):
		if not pd.isnull(row[x]) and isinstance(row[x], str):
			tokens = nltk.word_tokenize(row[x].replace('-', ' '))
			for y in range(len(tokens)):
				g = tokens[y]
				tokens[y] = tokens[y].upper()
				if tokens[y] in (name.upper() for name in bef):
					if index not in duplicates:
						if pasa.at[index,'Mapped System'] in msbm:
							yesno[index] = ['n/a', 'n/a']
						else:
							yesno[index] = ['n', 'n']
						duplicates[index] = [(0,pasa.at[index,'Mapped System'])]
					temp = ""
					for i in range(10):
						if y - 5 + i >= 0 and y - 5 + i < len(tokens):
							temp = temp + " " + tokens[y-5+i]
					duplicates[index].append((sorto[listo[x]], dic_b[tokens[y]] + " " + "{" + g + "}" + " " + listo[x] + ":: " + temp, dic_b[tokens[y]]))

#Analyze Mid
for index, row in data.iterrows():
	for x in range(len(row)):
		if not pd.isnull(row[x]) and isinstance(row[x], str):
			tokens = nltk.word_tokenize(row[x].replace('-', ' '))
			for y in range(len(tokens)):
				g = tokens[y]
				tokens[y] = tokens[y].upper()
				if tokens[y] in (name.upper() for name in mid):
					kis = True
					if tokens[y] == "KISAM":
						for i in range(10):
							if y - 5 + i >= 0 and y - 5 + i < len(tokens):
								if tokens[y-5+i].lower() == "incident":
									kis = False
					if tokens[y] == "TDS" and sorto[listo[x]] == 3:
						kis = False
					if kis:
						if index not in duplicates:
							if pasa.at[index,'Mapped System'] in msbm:
								yesno[index] = ['n/a', 'n/a']
							else:
								yesno[index] = ['n', 'n']
							duplicates[index] = [(0,pasa.at[index,'Mapped System'])]
						temp = ""
						for i in range(10):
							if y - 5 + i >= 0 and y - 5 + i < len(tokens):
								temp = temp + " " + tokens[y-5+i]
						duplicates[index].append((sorto[listo[x]], dic_m[tokens[y]] + " " + "{" + g + "}" + " " + listo[x] + ":: " + temp, dic_m[tokens[y]]))

#Analyze End
for index, row in data.iterrows():
	for x in range(len(row)):
		if not pd.isnull(row[x]) and isinstance(row[x], str):
			tokens = nltk.word_tokenize(row[x].replace('-', ' '))
			for y in range(len(tokens)):
				g = tokens[y]
				tokens[y] = tokens[y].upper()
				if tokens[y] in (name.upper() for name in end):
					kis = True
					if tokens[y] == "CUSTCOMM" and sorto[listo[x]] != 3:
						kis = False
					if tokens[y] == "TELECOM" and sorto[listo[x]] != 3:
						kis = False
					if kis:
						if index not in duplicates:
							if pasa.at[index,'Mapped System'] in msbm:
								yesno[index] = ['n/a', 'n/a']
							else:
								yesno[index] = ['n', 'n']
							duplicates[index] = [(0,pasa.at[index,'Mapped System'])]
						temp = ""
						for i in range(10):
							if y - 5 + i >= 0 and y - 5 + i < len(tokens):
								temp = temp + " " + tokens[y-5+i]
						if tokens[y] == "E2E":
							duplicates[index].append((100, dic_e[tokens[y]] + " "+ "{" + g + "}" + " " + listo[x] + ":: " + temp, dic_e[tokens[y]]))
						else:
							duplicates[index].append((sorto[listo[x]], dic_e[tokens[y]] + " "+ "{" + g + "}" + " " + listo[x] + ":: " + temp, dic_e[tokens[y]]))


#Analyze Special. Instead of looking at each token in each index/column. We look at each token from our special list of tokens and search for them within the substrings of each index/column. This allows us to find matches for tokens that include spaces.
for index, row in data.iterrows():
	for x in range(len(row)):
		if not pd.isnull(row[x]) and isinstance(row[x], str):
			for y in range(len(special)):
				z = row[x].find(special[y])
				if z != -1:
					if index not in duplicates:
						if pasa.at[index,'Mapped System'] in msbm:
							yesno[index] = ['n/a', 'n/a']
						else:
							yesno[index] = ['n', 'n']
						duplicates[index] = [(0,pasa.at[index,'Mapped System'])]
					temp = ""
					for i in range(80):
						if z - 40 + i >= 0 and z - 40 + i < len(row[x]):
							temp = temp + row[x][z-40+i]
					duplicates[index].append((sorto[listo[x]], dic_s[special[y]] + " "+ "{" + special[y] + "}" + " " + listo[x] + ":: " + temp, dic_s[special[y]]))
	#For when no tokens are found
	if index not in duplicates:
		if pasa.at[index,'Mapped System'] in msbm:
			yesno[index] = ['n/a', 'n/a']
		else:
			yesno[index] = ['n', 'n']
		duplicates[index] = [(0, pasa.at[index, 'Mapped System'])]
		duplicates[index].append((1, "", ""))

#List of systems we value least
lastpick = ["E2E-BSM", "Network", "SECURITY", "COTS"]

#Copy of duplicates
duplicates2 = {}
#Reorders certain matches based on lastpick
for x in duplicates.keys():
	duplicates[x] = sorted(duplicates[x], key = lambda k: k[0])
	if duplicates[x][1][2] in lastpick:
		for y in range(2,len(duplicates[x])):
			if duplicates[x][y][2] not in lastpick:
				temp = duplicates[x][1]
				duplicates[x][1] = duplicates[x][y]
				duplicates[x][y] = temp
	if duplicates[x][0][1] == duplicates[x][1][2]:
		yesno[x][0] = 'y'
	for y in range(len(duplicates[x]) - 2):
		if duplicates[x][0][1] == duplicates[x][y+2][2]:
			yesno[x][1] = 'y'
	duplicates2[x] = [lis[1] for lis in duplicates[x]] 

falsepos = {}
trueno = len(yesno)
for x in yesno:
	if yesno[x][0] == 'n/a':
		trueno = trueno - 1
#Sort by system
systemsort = {}
for x in duplicates.keys():
	if duplicates[x][0][1] not in systemsort:
		systemsort[duplicates[x][0][1]] = [x]
	else:
		systemsort[duplicates[x][0][1]].append(x)
	if duplicates[x][1][2] not in falsepos:
		if duplicates[x][1][2] != duplicates[x][0][1]:
			falsepos[duplicates[x][1][2]] = 1
	else:
		if duplicates[x][1][2] != duplicates[x][0][1]:
			falsepos[duplicates[x][1][2]] = falsepos[duplicates[x][1][2]] + 1


for x in systemsort.keys():
	truepos = 0
	falseneg = 0
	for y in systemsort[x]:
		if yesno[y][0] == 'y':
			truepos = truepos + 1
		elif yesno[y][0] == 'n':
			falseneg = falseneg + 1
	if x in falsepos:
		trueneg = trueno - truepos - falseneg - falsepos[x]
		systemsort[x] = [x, truepos, falseneg, falsepos[x], trueneg]
	else:
		systemsort[x] = [x, truepos, falseneg, 0, trueno - truepos - falseneg]
systemsort = sorted(systemsort.items(), key = lambda k: 0 if (k[1][1] == 0) else (k[1][1]/(k[1][1] + k[1][3])), reverse = True)
tempo = [["System Name", "True Pos", "False Neg", "False Pos", "True Neg"]]
for x in systemsort:
	tempo.append(x[1])
systemsort = pd.DataFrame(tempo)
#list that we're outputting to excel
temp = [["Incident ID", "Filter", "Has SRT", "On PSL?", "Match Found", "Duplicate Match Found", "Mapped System", "Found System", "Duplicates"]]
#correct matches
num = 0
#wrong matches with correct match found but not picked
numo = 0
#wrong match where the correct match was never found
numa = 0
#Creates output
for x in duplicates2.keys():
	temp.append([incid.at[x,'Incident ID'], incid.at[x,'Filter'], incid.at[x,'Has SRT'], incid.at[x, 'On PSL?'], yesno[x][0], yesno[x][1]] + duplicates2[x])
	if yesno[x][0] == 'y':
		num = num + 1
	if yesno[x][0] == 'n' and yesno[x][1] == 'y':
		numo = numo + 1
	if yesno[x][0] == 'n' and yesno[x][1] == 'n':
		numa = numa + 1
duplicates2 = pd.DataFrame(temp)
#print data
print(num)
print(numo)
print(numa)
#output to Excel
with pd.ExcelWriter('output.xlsx') as writer:  
	duplicates2.to_excel(writer, sheet_name = 'Output',index = False)
	systemsort.to_excel(writer, sheet_name = 'By System', index = False)