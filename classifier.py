import nltk
import pandas as pd
import numpy as np
from operator import itemgetter
import easygui

#choose file
path = easygui.fileopenbox()
#read excel to dataframe
services = pd.read_excel(path, sheet_name = 'System List', usecols = "A,K")
data = pd.read_excel(path, sheet_name = 'Combined', usecols = "A,B,C,H,I,J,K,T,X,Y,Z,AA,AB,AE,AJ,AP,BA,BB")

pas = []
for index, row in services.iterrows():
	if not pd.isnull(row[1]):
		pas.append(row[1])
pas.remove('Applications')

listo = ['Assignment Group Office','Assignment Group Division','Assignment Group','Title','Primary Affected Service','Subcategory','Area','Solution','Affected CI','Escalation Teams','PM Candidate Full Name','Problem Management Candidate','PM Candidate Requested By','Description','Parent Incident','Mapped System','Modified Assignment Group Office','Left 12 of Title']
dic = {}
duplicates = []
cols = [0] * len(listo)
for index, row in data.iterrows():
	for x in range(len(row)):
		if not pd.isnull(row[x]) and isinstance(row[x], str):
			tokens = nltk.word_tokenize(row[x])
			for y in tokens:
				if y in pas:
					cols[x] = cols[x] + 1
					if index not in dic:
						dic[index] = (x,y)
					elif y == dic[index][1]:
						dic[index] = (x,y)
					else:
						duplicates.append([index,y,dic[index][1]])


print(len(dic))
print(len(duplicates))
res = dict(zip(listo, cols))
print(res) 