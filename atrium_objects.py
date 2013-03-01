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
        month = ["", "Jan", "Feb", "March", "April", "May", "June", "July", "Aug", "Sept", "Oct", "Nov", "Dec"]
        self.exactDate = str(exactDate[0]) + " " + month[int(exactDate[1])] + " " + str(exactDate[2]) + " " + str(exactDate[3]) + ":" + str(exactDate[4])
        self.dictList = dictList
        self.numPeople = 0
        self.numGroups = 0
        self.dayDate = (int(exactDate[0]), int(exactDate[1]), int(exactDate[2]))
        self.dayDateString = str(exactDate[0]) + " " + month[int(exactDate[1])] + " " + str(exactDate[2])
        self.personList = []
        self.groupList = []
    
    def addPerson(self, person):
        self.personList.append(person)
        self.numPeople += 1
    
    def addGroup(self, group):
        self.groupList.append(group)
        self.numGroups += 1

    def addSmallChair(self, chair):
        self.chair = chair

    def addSofa(self, sofa):
        self.sofa = sofa

    def addSmallTable(self, smallTable):
        self.smallTable = smallTable

    def addLargeTable(self, largeTable):
        self.largeTable = largeTable
        
    def addPingPongTable(self, table):
        self.pingPongTable = table

    def __str__(self):
         return self.exactDate.toString()       

#knows number of people in it, and which people     
class Group:
    def __init__(self, dicty):
        self.minX = dicty["x"]
        self.minY = dicty["y"]
        self.maxX = dicty["x"] + dicty["width"]
        self.maxY = dicty["y"] + dicty["height"]
        self.numPeople = 0
        self.personList = []
        
    def addPerson(self, person):
        self.personList.append(person)
        self.numPeople += 1
        
    def isPersoninGroup(self, person):
        if person.midX >= self.minX and person.midX <= maxX and person.midY >= self.minY and person.midY <= self.maxY:
            return true
        else:
            return false

#knows center point of person
class Person:
    def __init__(self, dicty):
        self.midX = float(dicty["x"]) + float(dicty["width"]) / 2.0
        self.midY = float(dicty["y"]) + float(dicty["height"]) / 2.0


class Chair:
    def __init__(self, dicty):
        self.minX = dicty["x"]
        self.minY = dicty["y"]
        self.maxX = dicty["x"] + dicty["width"]
        self.maxY = dicty["y"] + dicty["height"]
        self.numPeople = 0
        self.personList = []

    def addPerson(self, person):
        self.personList.append(person)
        self.numPeople += 1
        
    def isPersoninGroup(self, person):
        if person.midX >= self.minX and person.midX <= maxX and person.midY >= self.minY and person.midY <= self.maxY:
            return true
        else:
            return false

3 = small chair
4 = sofa
5 = small table
6 = large table
7 = ping pong table
