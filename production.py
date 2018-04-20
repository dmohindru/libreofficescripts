# Author: Dhruv Mohindru
# This file contains macros to implement productions rated operation
# like generating summary, update master carton xml file and
# loading cost price from xml generated from Tally.
import sys
class CartonList:
  def __init__(self, startCartonNum):
    self.startCartonNum = startCartonNum
    self.endCartonNum = ""
  
  def setEndCartonNum(self, endCartonNum):
    self.endCartonNum = endCartonNum
  
  def getStartCartonNum(self):
    return self.startCartonNum
  
  def getEndCartonNum(self):
    return self.endCartonNum


class CycleSpecs:
  
  def __init__(self, specList):
    self.specsList = specList
    self.count = 1
    self.cartonList = []
  
  def getSpecs(self):
    #self.specsList.sort()
    return self.specsList
  
  def getCartonsList(self):
    return self.cartonList
    
  def getCartonCount(self):
    return self.count
  
  def incrementCount(self):
    self.count += 1
  
  def isSpecsSame(self, specList):
    specLen = len(specList)
    i = 0
    for e in self.specsList:
      if e not in specList:
        return False
      i += 1
    if i != specLen:
      return False
    return True
  
  def setCartonList(self, cartons):
    # need to work on this function more
    self.cartonList.append(cartons)
  
  def getLastCartonObj(self):
    return self.cartonList[-1]

def testSpecsList():
  list1 = ["item2", "item1", "item4", "item3"]
  list2 = ["item1", "item2", "item3", "item4", "item5"]
  list3 = ["item1", "item3"]
  list4 = ["item6", "item1", "item3", "item4"]
  # Test constructor
  specs1 = CycleSpecs(list1)
  # Test getter and setter function
  print(specs1.getSpecs())
  print(specs1.getCartonCount())
  specs1.incrementCount()
  specs1.incrementCount()
  print(specs1.getCartonCount())
  # Test isSpecsSame function
  print(specs1.isSpecsSame(list2))
  print(specs1.isSpecsSame(list1))
  print(specs1.isSpecsSame(list3))
  print(specs1.isSpecsSame(list4))
  
  

class CycleModel:
  'Class for Cycle Model'
  
  # Class variable to hold total number of different model in sheet
  totalModels = 0
  
  def __init__(self, model):
    self.modelName = model # set name of model
    self.specsList = []     # set specs list to empty list
    CycleModel.totalModels += 1 # increment total model variable on creation of new model
  
  def getModelName(self):
    return self.modelName
  
  def getSpecsList(self):
    return self.specsList
  
  def getTotalModels(self):
    return CycleModel.totalModels
  
  def setSpecsList(self, specs):
    self.specsList.append(specs)
  
  def setModelName(self, name):
    self.modelName = name
  
  def __lt__(self, other):
    return self.modelName < other.modelName
  
# Temp function to check for funcanality
def testCycleModel():
  specsList = ["BT", "Seat 6000", "Twisted Spoke"]
  modelList = []
  cycle = CycleModel("Dragon")
  cycle.setSpecsList(specsList)
  # Test getter methods
  print(cycle.getModelName())
  print(cycle.getSpecsList())
  print(cycle.getTotalModels())

  # Test setter methods
  cycle.setModelName("Don BXM")
  cycle.setSpecsList(["Eva", "4 Pcs"])
  print(cycle.getModelName())
  print(cycle.getSpecsList())
  print(cycle.getTotalModels())
  #modelList.append(cycle)
  
  cycle = CycleModel("Dragon")
  cycle.setSpecsList(specsList)
  print(cycle.getTotalModels())
  
  
  
  

def Initalize():
#get the doc from the scripting context which is made available to all scripts
  desktop = XSCRIPTCONTEXT.getDesktop()
  model = desktop.getCurrentComponent()
  return model
    
# function to generate summary for production of whole month
def generateSummary(*args):
  # Cell column from where entries start
  startCellCol = "D"
  # Cell row from where entries start
  startCellRow = 4
  # Concatinate to generate a format of form A2
  cellNo = startCellCol + str(startCellRow)
  DATA_SHEET = 0
  modelList = [] # List of all models in sheet
  numOfCartons = 0
   
  # Initalize and get required context
  model = Initalize()
  sheet = model.Sheets.getByIndex(DATA_SHEET);
  
  currentModel = sheet.getCellRangeByName(cellNo).String
  previousModel = None
  currentSpecs = readCurrentSpecs(sheet, startCellRow)
  previousSpecs = None
  
  #read specs test
  #temp = readCurrentSpecs(sheet, 4)
  #col = "O"
  #row = 4
  #for e in temp:
  #  sheet.getCellRangeByName(col+str(row)).String = e
  #  row += 1
  ############
  
  while currentModel != "":
    
    currentModelObj = getCurrentModelObj(modelList, currentModel)
    # Read current carton number
    currentCartonNum = readCartonNum(sheet, startCellRow)
    # Read specs for current row
    #currentSpecs = readCurrentSpecs(sheet, startCellRow)
    
    if previousModel != currentModel:
      if currentModelObj != None: # Model is present
        specsFound = False
        for e in currentModelObj.getSpecsList():
          if e.isSpecsSame(currentSpecs): # Specs are present
            e.incrementCount()
            specsFound = True
            # create new Carton number object with current carton number
            cartonNumObj = CartonList(currentCartonNum)
            # set the new carton number object to specsList carton number list
            e.setCartonList(cartonNumObj)
            break
        if specsFound == False:
          newSpecs = CycleSpecs(currentSpecs)
          # create new Carton number object with current carton number
          cartonNumObj = CartonList(currentCartonNum)
          # set the new carton number object to specsList carton number list
          newSpecs.setCartonList(cartonNumObj)
          currentModelObj.setSpecsList(newSpecs)
            
      else: # Model is not present
        # create new Carton number object with current carton number
        cartonNumObj = CartonList(currentCartonNum)
        # create new specs
        newSpecs = CycleSpecs(currentSpecs)
        # set the new carton number object to specsList carton number list
        newSpecs.setCartonList(cartonNumObj)
        # create new model
        newCycleModel = CycleModel(currentModel)
        # append specs list
        newCycleModel.setSpecsList(newSpecs)
        modelList.append(newCycleModel)
    elif previousSpecs != currentSpecs: # current specs are different
      specsFound = False
      for e in currentModelObj.getSpecsList():
        if e.isSpecsSame(currentSpecs): # Specs are present
          e.incrementCount()
          # create new Carton number object with current carton number
          cartonNumObj = CartonList(currentCartonNum)
          # set the new carton number object to specsList carton number list
          e.setCartonList(cartonNumObj)
          specsFound = True
          break
      if specsFound == False:
        newSpecs = CycleSpecs(currentSpecs)
        currentModelObj.setSpecsList(newSpecs)
        # create new Carton number object with current carton number
        cartonNumObj = CartonList(currentCartonNum)
        # set the new carton number object to specsList carton number list
        newSpecs.setCartonList(cartonNumObj)
        
    else: # specs are same increment count
      # find specs and increment it in currentModel
      for e in currentModelObj.getSpecsList():
        if e.isSpecsSame(currentSpecs): # Specs are present
          e.incrementCount()
          # get the last element of carton number list from specs list
          lastCartonListObj = e.getLastCartonObj()
          # record current carton number as end carton number of list
          lastCartonListObj.setEndCartonNum(currentCartonNum)
          break
      
    numOfCartons += 1
    startCellRow += 1
    cellNo = startCellCol + str(startCellRow)
    previousModel = currentModel
    previousSpecs = currentSpecs
    currentModel  = sheet.getCellRangeByName(cellNo).String
    #Read specs for current row
    currentSpecs = readCurrentSpecs(sheet, startCellRow)
  
  printSummary(modelList, numOfCartons)
  
  return None

def printSummary(modelList, numOfCartons):
  'This function prints summary of production'
  # Cell column from where entries start
  startCellCol = "B"
  # Cell row from where entries start
  startCellRow = 2
  # Concatinate to generate a format of form A2
  cellNo = startCellCol + str(startCellRow)
  SUMMARY_SHEET = 1
  #modelList = [] # List of all models in sheet
  #numOfCartons = 0
   
  # Initalize and get required context
  model = Initalize()
  sheet = model.Sheets.getByIndex(SUMMARY_SHEET);
  modelList.sort()
  for model in modelList:
    cellNo = startCellCol + str(startCellRow)
    sheet.getCellRangeByName(cellNo).String = model.getModelName()
    startCellRow += 1
    # get specs list
    specsList = model.getSpecsList()
    # generate string of form (item1, item2, item3 ...etc)
    
    for specs in specsList:
      # Increment Row number
      #startCellRow += 1
      specsStr = "("
      singleSpecs = specs.getSpecs()
      numOfItems = len(singleSpecs)
      # Generate a single specs string 
      for e in singleSpecs:
        if numOfItems > 1:
          specsStr += e + ", "
        else:
          specsStr += e
        numOfItems -= 1
      specsStr += ")"
      # Increment Row number
      #startCellRow += 1
      # Print count of specs at column C
      cellNo = "C" + str(startCellRow)
      sheet.getCellRangeByName(cellNo).String = specs.getCartonCount()
      # Print specs string at column D
      cellNo = "D" + str(startCellRow)
      sheet.getCellRangeByName(cellNo).String = specsStr
      # Print carton number list at column E
      cartonNumString = "("
      cartonNumList = specs.getCartonsList()
      numOfCartonObj = len(cartonNumList)
      # Generate carton numbers string for single specs
      for cartonNum in cartonNumList:
        cartonNumString += cartonNum.getStartCartonNum()
        endCartonNum = cartonNum.getEndCartonNum()
        if endCartonNum != "":
          cartonNumString += ":" + endCartonNum
        if numOfCartonObj > 1:
          cartonNumString += ", "
        numOfCartonObj -= 1
      cartonNumString += ")"
      cellNo = "E" + str(startCellRow)
      sheet.getCellRangeByName(cellNo).String = cartonNumString
      startCellRow += 1
  
  startCellRow += 1
  cellNo = startCellCol + str(startCellRow)
  sheet.getCellRangeByName(cellNo).String = "Total: " + str(numOfCartons)
  

def getCurrentModelObj(modelList, currentModel):
  'This model reterives model obj from list of models'
  for model in modelList:
    if model.getModelName() == currentModel:
      return model
  return None

def readCurrentSpecs(sheet, row):
  'This function reads specs for the current row in sheet'
  maxSpecs = 6
  specsCounter = 1
  specsStartCol = "G"
  cellNo = specsStartCol + str(row)
  specs = []
  component = sheet.getCellRangeByName(cellNo).String
  while component != "" and specsCounter <= 6:
    specs.append(component)
    specsStartCol = chr(ord(specsStartCol) + 1)
    specsCounter += 1
    cellNo = specsStartCol + str(row)
    component = sheet.getCellRangeByName(cellNo).String
  return specs

def readCartonNum(sheet, row):
  'This function reads and returns the current carton number'
  cartonNumCol = "A"
  cellNo = cartonNumCol + str(row)
  return sheet.getCellRangeByName(cellNo).String
  

  

