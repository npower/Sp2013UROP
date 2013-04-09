#gathers info from photos across a full day
class Day:
    def __init__(self, day):
        self.day = day
        self.photoList = []
        
    def addPhoto(self, photo):
        self.photoList.append(photo)

#stores info of a single photo      
class Photo:
    def __init__(self, exactDate, dictList):
        self.exactDate = exactDate
        self.dictList = dictList
        self.numPeople = 0
        self.numGroups = 0        
        self.dayDate = exactDate.date()
        self.personList = []
        self.groupList = []
        self.chairList = []
        self.sofaList = []
        self.smallTableList = []
        self.largeTableList = []
        #self.pingPongTable
    
    def addPerson(self, person):
        self.personList.append(person)
        self.numPeople = len(set(self.personList))
    
    def addGroup(self, group):
        self.groupList.append(group)
        self.numGroups = len(set(self.groupList))

    def addChair(self, chair):
        self.chairList.append(chair)

    def addSofa(self, sofa):
        self.sofaList.append(sofa)

    def addSmallTable(self, smallTable):
        self.smallTableList.append(smallTable)

    def addLargeTable(self, largeTable):
        self.largeTableList.append(largeTable)
        
##    def addPingPongTable(self, table):
##        self.pingPongTable = table

    def isOutlier(self):
        x = self.exactDate
        if x >= datetime(2013, 2, 1, 9, 30) and x <= datetime(2013, 2, 2, 15, 45):
            return True
        elif x >= datetime(2013, 2, 21, 13, 0) and x <= datetime(2013, 2, 21, 14, 45):
            return True
        elif x >= datetime(2013, 2, 22, 12, 00) and x <= datetime(2013, 2, 25, 10, 0):
            return True
        elif x >= datetime(2013, 2, 28, 10, 0) and x <= datetime(2013, 2, 28, 2, 40):
            return True
        elif x >= datetime(2013, 2, 28, 8, 0) and x <= datetime(2013, 2, 28, 11, 0):
            return True
        elif x >= datetime(2013, 3, 6, 9, 15) and x <= datetime(2013, 3, 7, 14, 0):
            return True
        elif x >= datetime(2013, 3, 8, 16, 0) and x <= datetime(2013, 3, 8, 17, 0):
            return True
        elif x >= datetime(2013, 3, 9, 18, 0) and x <= datetime(2013, 3, 10, 13, 45):
            return True
        elif x >= datetime(2013, 3, 19, 16, 0) and x <= datetime(2013, 3, 20, 14, 0):
            return True
        elif x >= datetime(2013, 3, 28, 10, 0) and x <= datetime(2013, 3, 29, 12, 0):
            return True
        elif x >= datetime(2013, 4, 1, 12, 0) and x <= datetime(2013, 4, 1, 13, 0):
            return True
        else:
            return false        
        
    def __str__(self):
         return self.exactDate.toString()

#knows number of people in it, and which people     
class Group:
    def __init__(self, dicty):
        self.minX = int(dicty["x"])
        self.minY = int(dicty["y"])
        self.maxX = int(dicty["x"]) + int(dicty["width"])
        self.maxY = int(dicty["y"]) + int(dicty["height"])
        self.midX = float(dicty["x"]) + float(dicty["width"]) / 2.0
        self.midY = float(dicty["y"]) + float(dicty["height"]) / 2.0
        self.numPeople = 0
        self.personList = []
        
    def addPerson(self, person):
        self.personList.append(person)
        self.numPeople = len(set(self.personList))
        
    def isInGroup(self, thing):
        if thing.midX >= self.minX and thing.midX <= self.maxX and thing.midY >= self.minY and thing.midY <= self.maxY:
            return True
        else:
            return False
        

#knows center point of person
class Person:
    def __init__(self, dicty):
        self.dicty = dicty
        self.midX = float(dicty["x"]) + float(dicty["width"]) / 2.0
        self.midY = float(dicty["y"]) + float(dicty["height"]) / 2.0
        self.minX = int(dicty["x"])
        self.minY = int(dicty["y"])
        self.maxX = int(dicty["x"]) + int(dicty["width"])
        self.maxY = int(dicty["y"]) + int(dicty["height"])

    def __str__(self):
        return str([str(self.dicty["x"]), str(self.dicty["y"])])
        

class Chair:
    def __init__(self, dicty):
        self.dicty = dicty
        self.minX = int(dicty["x"])
        self.minY = int(dicty["y"])
        self.maxX = int(dicty["x"]) + int(dicty["width"])
        self.maxY = int(dicty["y"]) + int(dicty["height"])
        self.numPeople = 0
        self.numGroups = 0
        self.personList = []
        self.groupList = []

    def addPerson(self, person):
        self.personList.append(person)
        self.numPeople = len(set(self.personList))

    def addGroup(self, group):
        self.groupList.append(group)
        self.numGroups = len(set(self.groupList))
        
    def usingChair(self, person):
        if person.midX >= self.minX and person.midX <= self.maxX and person.midY >= self.minY and person.midY <= self.maxY:
            return True
        else:
            return False

    def __str__(self):
        return str([str(self.dicty["x"]), str(self.dicty["y"])])

class Sofa:
    def __init__(self, dicty):
        self.minX = int(dicty["x"])
        self.minY = int(dicty["y"])
        self.maxX = int(dicty["x"]) + int(dicty["width"])
        self.maxY = int(dicty["y"]) + int(dicty["height"])
        self.numPeople = 0
        self.numGroups = 0
        self.personList = []
        self.groupList = []

    def addPerson(self, person):
        self.personList.append(person)
        self.numPeople = len(set(self.personList))

    def addGroup(self, group):
        self.groupList.append(group)
        self.numGroups = len(set(self.groupList))
        
    def usingSofa(self, thing):
        if thing.midX >= self.minX and thing.midX <= self.maxX and thing.midY >= self.minY and thing.midY <= self.maxY:
            return True
        else:
            return False

class Small_Table:
    def __init__(self, dicty):
        self.minX = int(dicty["x"])
        self.minY = int(dicty["y"])
        self.maxX = int(dicty["x"]) + int(dicty["width"])
        self.maxY = int(dicty["y"]) + int(dicty["height"])
        self.numPeople = 0
        self.numGroups = 0
        self.personList = []
        self.groupList = []

    def addPerson(self, person):
        self.personList.append(person)
        self.numPeople = len(set(self.personList))

    def addGroup(self, group):
        self.groupList.append(group)
        self.numGroups = len(set(self.groupList))
        
    def usingSmallTable(self, thing):
        if thing.midX >= self.minX and thing.midX <= self.maxX and thing.midY >= self.minY and thing.midY <= self.maxY:
            return True
        else:
            return False

class Large_Table:
    def __init__(self, dicty):
        self.minX = int(dicty["x"])
        self.minY = int(dicty["y"])
        self.maxX = int(dicty["x"]) + int(dicty["width"])
        self.maxY = int(dicty["y"]) + int(dicty["height"])
        self.midX = float(dicty["x"]) + float(dicty["width"]) / 2.0
        self.midY = float(dicty["y"]) + float(dicty["height"]) / 2.0
        self.numPeople = 0
        self.numGroups = 0
        self.personList = []
        self.groupList = []

    def addPerson(self, person):
        self.personList.append(person)
        self.numPeople = len(set(self.personList))
        
    def addGroup(self, group):
        self.groupList.append(group)
        self.numGroups = len(set(self.groupList))
        
    def usingLargeTable(self, thing):
        if thing.midX >= self.minX and thing.midX <= self.maxX and thing.midY >= self.minY and thing.midY <= self.maxY:
            return True
        elif self.midX >= thing.minX and self.midX <= thing.maxX and self.midY >= thing.minY and self.midY <= thing.maxY:
            return True
        else:
            return False
        

##3 = small chair
##4 = sofa
##5 = small table
##6 = large table
##7 = ping pong table
