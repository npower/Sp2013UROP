##Atrium Photo analysis
import os


# input: directory path(string)
# goes through all the files in a directory and calls parseFile on them
# joins the outputs of all the parseFile calls and returns it
def readFiles(path):
	allDictList = []
	for file in os.listdir(path):	
		for dict in parseFile(path + "\\" + file): # file needs to be the complete file path
			allDictList.append(dict)	
			
	return allDictList

	
# input: a filename (string)
# opens the file, goes through all the lines and extracts the fields
# returns an unique list of dictionaries (one for each line)	
def parseFile(filePath):
	f = open(filePath) 
	linesList = f.readlines() 
	f.close()
	
	noDuplicates = list(set(linesList))
	
	fileDictList = []
	for line in noDuplicates:
		fileDictList.append(tokenize(line))
		
	return fileDictList

	
separators = [',', '_', '.']
# Converts string into a list of tokens (strings)
# save tokens in a dictionary, and returns it
def tokenize(string):
    tokensList = []  
    token = ''    #keeps track of a single element
    for x in string:
        if x in separators:
            if token != '':  #add nonempty token when a separator is reached, then reset token
                tokensList.append(token) 
                token = ''
        else:   
            token += x
	
    objectDict = {'ID':tokensList[0], 'x':tokensList[1], 'y':tokensList[2], 'width':tokensList[3], \
        'height':tokensList[4], 'year':tokensList[5], 'month':tokensList[6], 'day':tokensList[7], \
        'hour':tokensList[8],	'minute':tokensList[9], 'second':tokensList[10]}
			
    return objectDict
	
	
# this is a generic function to help with calculating statistics on the dataset
# specific functions could be written instead
# input: data: list of dictionaries, fun: a lambda function that takes in a dictionary and outputs a number
# returns: a list of numbers that were obtained by applying fun on each element in data
def calculateStat(data, fun):
	
	pass

	
# generates an Excel spreadsheet
# input: a list of lists or dictionaries to put in the spreadsheet
# each first-level list is a sheet in Excel
# example usage: outputToExcel(calculateStat(allData, *function that calculates average of X*), calculateStat(allData, *function that calculates average of Y*))
def outputToExel(data):
	pass
	


