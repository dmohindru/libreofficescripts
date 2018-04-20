# Author: Dhruv Mohindru
#this file implements updating of master carton xml and sqlite database file to store monthly production cost of cycle
import sys
from xml.dom.minidom import parse
import xml.dom.minidom

#shortcut used in creating file dialogs
createUnoService = (
        XSCRIPTCONTEXT
        .getComponentContext()
        .getServiceManager()
        .createInstance 
                    )
# Function to get model object to access sheet in a document
def Initalize():
  #get the doc from the scripting context which is made available to all scripts
  desktop = XSCRIPTCONTEXT.getDesktop()
  model = desktop.getCurrentComponent()
  return model

# Function to update master carton xml for monthly production cost
def updateMasterCartonXML(*args):
  PRODUCTION_SHEET = 0
  # Initalize and get required context
  model = Initalize()
  sheet = model.Sheets.getByIndex(PRODUCTION_SHEET);
  sheet.getCellRangeByName("A2").String = "I am update XML button"
  return None
  
# Function to update sqlite database file for monthly production cost
def updateMasterDBFile(*args):
  PRODUCTION_SHEET = 0
  # Initalize and get required context
  model = Initalize()
  sheet = model.Sheets.getByIndex(PRODUCTION_SHEET);
  sheet.getCellRangeByName("A2").String = "I am update DB button"
  return None


