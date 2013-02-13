##Atrium Photo analysis

def readFiles(directory):
	# input: directory (string)
	# goes through all the files in a directory and calls parseFile on them
	# joins the outputs of all the parseFile calls and returns it
	pass

def parseFile(fileName):
	# input: a filename (string)
	# opens the file, goes through all the lines and extracts the fields
	# returns an unique list of dictionaries (one for each line)
	pass

def calculateStat(data, fun):
	# this is a generic function to help with calculating statistics on the dataset
	# specific functions could be written instead
	# input: data: list of dictionaries, fun: a lambda function that takes in a dictionary and outputs a number
	# returns: a list of numbers that were obtained by applying fun on each element in data
	pass

def outputToExel(data):
	# generates an Excel spreadsheet
	# input: a list of lists or dictionaries to put in the spreadsheet
	# each first-level list is a sheet in Excel
	# example usage: outputToExcel(calculateStat(allData, *function that calculates average of X*), calculateStat(allData, *function that calculates average of Y*))
	pass
	

# characters that are single-character tokens
separators = [',']

# Convert strings into a list of tokens (strings)
def tokenize(string):
    tokensList = []  #list of elements in the string, in order
    token = ''    #keeps track of a single element
    for x in string:
        if x in separators:
            if token != '':  #add the nonempty token when it reaches a separator, then resets token
                tokensList.append(token) 
                token = ''
        else:   
            token += x
    if token != '':  #adds final token
        tokensList.append(token)
        token = ''
    return tokensList