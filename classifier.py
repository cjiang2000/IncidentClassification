import nltk
import pandas as pd
import numpy as np
from operator import itemgetter
import easygui

#choose file
path = easygui.fileopenbox()
#read excel to dataframe
services = pd.read_excel("SystemList.xlsx", usecols = "N,O,P,Q,R,S,T,U,V,W")
data = pd.read_excel(path, sheet_name = 'Combined', usecols = "X,C,T,H,R,Y,AE")
pasa = pd.read_excel(path,sheet_name = 'Combined', usecols = "AP")
incid = pd.read_excel(path,sheet_name = 'Combined', usecols = "E,BN,BO")

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
msbm = ["Applications", "SECURITY", "Mainframe", "Data / Database", "Server Infrastructure", "Service Desk"]
for index, row in services.iterrows():
	if not pd.isnull(row["Before"]):
		bef.append(row["Before"])
		dic_b[row["Before"]] = row["Match_b"]
	if not pd.isnull(row["Middle"]):
		mid.append(row["Middle"])
		dic_m[row["Middle"]] = row["Match_m"]
	if not pd.isnull(row["End"]):
		end.append(row["End"])
		dic_e[row["End"]] = row["Match_e"]
	if not pd.isnull(row["Special"]):
		special.append(row["Special"])
		dic_s[row["Special"]] = row["Match_s"]
	if not pd.isnull(row["mef"]):
		mef.append(row["mef"])
		dic_mef[row["mef"]] = row["Match_mef"]

count = 0
counto = [0] * data.shape[0]
listo = list(data.columns)
sorto = {"Affected CI": 1, "Assignment Group": 2, "Solution": 4, "Title": 5, "Update Action": 6, "Escalation Teams": 3, "Description": 7}

dic = {}
duplicates = {}
yesno = {}
cols = [0] * len(listo)

for index, row in data.iterrows():
	for x in range(len(row)):
		if not pd.isnull(row[x]) and isinstance(row[x], str):
			tokens = nltk.word_tokenize(row[x].replace('-', ' '))
			for y in range(len(tokens)):
				if tokens[y] in mef:
					if index not in duplicates:
						if pasa.at[index,'Mapped System'] in msbm:
							yesno[index] = ['n/a', 'n/a']
						else:
							yesno[index] = ['n', 'n']
						duplicates[index] = [(0,pasa.at[index,'Mapped System'])]
	#					duplicates[index].append(dic[index][1] + " " + listo[x])
					if pasa.at[index,'Mapped System'] == dic_mef[tokens[y]]:
						yesno[index][1] = 'y'
					temp = ""
					for i in range(10):
						if y - 5 + i >= 0 and y - 5 + i < len(tokens):
							temp = temp + " " + tokens[y-5+i]
					duplicates[index].append((sorto[listo[x]], dic_mef[tokens[y]] + " " + "{" + tokens[y] + "}" + " " + listo[x] + ":: " + temp, dic_mef[tokens[y]]))
					if dic_mef[tokens[y]] == pasa.at[index, 'Mapped System']:
						counto[index] = 1

for index, row in data.iterrows():
	for x in range(len(row)):
		if not pd.isnull(row[x]) and isinstance(row[x], str):
			tokens = nltk.word_tokenize(row[x].replace('-', ' '))
			for y in range(len(tokens)):
				if tokens[y] in bef:
					if index not in duplicates:
						if pasa.at[index,'Mapped System'] in msbm:
							yesno[index] = ['n/a', 'n/a']
						else:
							yesno[index] = ['n', 'n']
						duplicates[index] = [(0,pasa.at[index,'Mapped System'])]
	#					duplicates[index].append(dic[index][1] + " " + listo[x])
					if pasa.at[index,'Mapped System'] == dic_b[tokens[y]]:
						yesno[index][1] = 'y'
					temp = ""
					for i in range(10):
						if y - 5 + i >= 0 and y - 5 + i < len(tokens):
							temp = temp + " " + tokens[y-5+i]
					duplicates[index].append((sorto[listo[x]], dic_b[tokens[y]] + " " + "{" + tokens[y] + "}" + " " + listo[x] + ":: " + temp, dic_b[tokens[y]]))
					if dic_b[tokens[y]] == pasa.at[index, 'Mapped System']:
						counto[index] = 1

for index, row in data.iterrows():
	if index not in dic:
		for x in range(len(row)):
			if not pd.isnull(row[x]) and isinstance(row[x], str):
				tokens = nltk.word_tokenize(row[x].replace('-', ' '))
				for y in range(len(tokens)):
					if tokens[y] in mid:
						if index not in duplicates:
							if pasa.at[index,'Mapped System'] in msbm:
								yesno[index] = ['n/a', 'n/a']
							else:
								yesno[index] = ['n', 'n']
							duplicates[index] = [(0,pasa.at[index,'Mapped System'])]
					#		duplicates[index].append(dic[index][1] + " " + listo[x])
						if pasa.at[index,'Mapped System'] == dic_m[tokens[y]]:
							yesno[index][1] = 'y'
						temp = ""
						for i in range(10):
							if y - 5 + i >= 0 and y - 5 + i < len(tokens):
								temp = temp + " " + tokens[y-5+i]
						duplicates[index].append((sorto[listo[x]], dic_m[tokens[y]] + " " + "{" + tokens[y] + "}" + " " + listo[x] + ":: " + temp, dic_m[tokens[y]]))
						if dic_m[tokens[y]] == pasa.at[index, 'Mapped System']:
							counto[index] = 1


for index, row in data.iterrows():
	if index not in dic:
		for x in range(len(row)):
			if not pd.isnull(row[x]) and isinstance(row[x], str):
				tokens = nltk.word_tokenize(row[x].replace('-', ' '))
				for y in range(len(tokens)):
					if tokens[y] in end:
						if index not in duplicates:
							if pasa.at[index,'Mapped System'] in msbm:
								yesno[index] = ['n/a', 'n/a']
							else:
								yesno[index] = ['n', 'n']
							duplicates[index] = [(0,pasa.at[index,'Mapped System'])]
				#			duplicates[index].append(dic[index][1] + " " + listo[x])
						if pasa.at[index,'Mapped System'] == dic_e[tokens[y]]:
							yesno[index][1] = 'y'
						temp = ""
						for i in range(10):
							if y - 5 + i >= 0 and y - 5 + i < len(tokens):
								temp = temp + " " + tokens[y-5+i]
						duplicates[index].append((sorto[listo[x]], dic_e[tokens[y]] + " "+ "{" + tokens[y] + "}" + " " + listo[x] + ":: " + temp, dic_e[tokens[y]]))
						if dic_e[tokens[y]] == pasa.at[index, 'Mapped System']:
							counto[index] = 1


for index, row in data.iterrows():
	if index not in dic:
		for x in range(len(row)):
			if not pd.isnull(row[x]) and isinstance(row[x], str):
				for y in range(len(special)):
					if special[y] in row[x]:
						if index not in duplicates:
							if pasa.at[index,'Mapped System'] in msbm:
								yesno[index] = ['n/a', 'n/a']
							else:
								yesno[index] = ['n', 'n']
							duplicates[index] = [(0,pasa.at[index,'Mapped System'])]
						if pasa.at[index,'Mapped System'] == dic_s[special[y]]:
							yesno[index][1] = 'y'
						temp = ""
						for i in range(80):
							if y - 40 + i >= 0 and y - 40 + i < len(row[x]):
								temp = temp + row[x][y-40+i]
						duplicates[index].append((sorto[listo[x]], dic_s[special[y]] + " "+ "{" + special[y] + "}" + " " + listo[x] + ":: " + temp, dic_s[special[y]]))
						if dic_s[special[y]] == pasa.at[index, 'Mapped System']:
							counto[index] = 1

res = dict(zip(listo, cols))
for x in duplicates.keys():
	duplicates[x] = sorted(duplicates[x], key = lambda k: k[0])
	if duplicates[x][0][1] == duplicates[x][1][2]:
		yesno[x][0] = 'y'
	for y in range(len(duplicates[x]) - 2):
		if duplicates[x][0][1] == duplicates[x][y+2][2]:
			yesno[x][1] = 'y'
	duplicates[x] = [lis[1] for lis in duplicates[x]] 

temp = [["Incident ID", "Filter", "Has SRT", "Match Found", "Duplicate Match Found", "Mapped System", "Found System", "Duplicates"]]
num = 0
numo = 0
numa = 0
for x in duplicates.keys():
	temp.append([incid.at[x,'Incident ID'], incid.at[x,'Filter'], incid.at[x,'Has SRT'], yesno[x][0], yesno[x][1]] + duplicates[x])
	if yesno[x][0] == 'y':
		num = num + 1
	if yesno[x][0] == 'n' and yesno[x][1] == 'y':
		numo = numo + 1
	if yesno[x][0] == 'n' and yesno[x][1] == 'n':
		numa = numa + 1
duplicates = pd.DataFrame(temp)

print(num)
print(numo)
print(numa)
with pd.ExcelWriter('output.xlsx') as writer:  
	duplicates.to_excel(writer, sheet_name = 'Dups',index = False)