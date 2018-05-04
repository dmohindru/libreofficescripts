# Author: Dhruv Mohindru
#this file implements updating of master carton xml and sqlite database file to store monthly production cost of cycle
import sys
from xml.dom.minidom import parse
import xml.dom.minidom
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL

from shutil import copyfile

#shortcut used in creating file dialogs
createUnoService = (
        XSCRIPTCONTEXT
        .getComponentContext()
        .getServiceManager()
        .createInstance 
                    )
# Function to get model object to access sheet in a document
def initalize():
  #get the doc from the scripting context which is made available to all scripts
  desktop = XSCRIPTCONTEXT.getDesktop()
  model = desktop.getCurrentComponent()
  return model
  
def errorBox(message, title):
  '''
  Function to invoke error dialog box 
  '''
  doc = XSCRIPTCONTEXT.getDocument()
  parentwin = doc.CurrentController.Frame.ContainerWindow
  box = parentwin.getToolkit().createMessageBox(parentwin, ERRORBOX,  BUTTONS_OK, title, message)
  result = box.execute()

def infoBox(message, title):
  '''
  Function to invoke error dialog box 
  '''
  doc = XSCRIPTCONTEXT.getDocument()
  parentwin = doc.CurrentController.Frame.ContainerWindow
  box = parentwin.getToolkit().createMessageBox(parentwin, INFOBOX,  BUTTONS_OK, title, message)
  result = box.execute()

def FilePicker(path=None, mode=1):
  '''
  Function to invoke file picker 
  '''
  filepicker = createUnoService("com.sun.star.ui.dialogs.FilePicker")
  filepicker.initialize( ( 0,) )
  filepicker.setMultiSelectionMode(False)
  filepicker.appendFilter("XML files (*.xml)", "*.xml")
  filepicker.setCurrentFilter("XML files (*.xml)")
  filepicker.execute()
  return filepicker.getFiles()[0]


'''
Structure of XML
<production>
  <month id="Apr17">
    <carton id="1">
      <modelno>201</model>
      <model>26X1.75 DRAGON</model>
      <pv>PV-28</pv>
      <price>2200.18</price>
    </carton>
    <carton id="2">
      <modelno>201</model>
      <model>26X1.75 DRAGON</model>
      <pv>PV-28</pv>
      <price>2200.18</price>
    </carton>
    ......
  <month>
  <month id="May17">
  .....
  </month>
  .....
</production>
Pseudo code for updading master carton xml file
  Read month from sheet
  if month not set
    show error message and return
  show file selection dialog to select master carton xml file
  opne master carton xml file
  make a backup copy of master carton xml file
  get list of month elements
  make <month> element and set it attribute id = month just read
  while no more entries in sheet
    read carton number column
    create <carton> element and set its attribute id = carton number
    read model number
    create <modelno> element and set its data to model number
    append as child element to <carton> element
    read model name
    create <model> element and set its data to model name
    append as child element to <carton> element
    read pv number
    create <pv> element and set its data to pv number
    append as child element to <carton> element
    read production price
    create <price> element and set its data to production price
    append as child element to <carton> element
    append as child element to <month> element
  append <month> element to list of month elements
  save xml
  display dialog notifing xml has been updated
'''
# Function to update master carton xml for monthly production cost
def updateMasterCartonXML(*args):
  PRODUCTION_SHEET = 0
  MONTH_CELLNO = "P1"
  CARTON_NO_COL = "A"
  CARTON_NO_ROW_START = 4
  MODEL_NO_COL = "C"
  MODEL_NAME_COL = "D"
  PV_COL = "O"
  PRICE_COL = "P"
  # Initalize and get required context
  model = initalize()
  sheet = model.Sheets.getByIndex(PRODUCTION_SHEET);
  if sheet.getCellRangeByName(MONTH_CELLNO).String == "":
    errorBox("Month not set", "Error")
    return None
  # Invoke File Picker to select MasterCartonXML.xml file
  filePath = FilePicker()[7:]
  # Return if no file is selected
  if not filePath:
    return None
  #Replace a single space in folder name Kingston Accounts which is returned as %20 by FilePicker
  filePath = filePath.replace('%20', ' ')
  # Show file path of selected file - TEMP STUFF
  #sheet.getCellRangeByName("A2").String = filePath
  #make a backup copy of master carton xml file
  newFilePath = filePath[:-4] + "-Backup" + filePath[-4:]
  #sheet.getCellRangeByName("B2").String = newFilePath
  copyfile(filePath, newFilePath)
  #get list of month elements
  DOMTree = xml.dom.minidom.parse(filePath)
  doc = xml.dom.minidom.Document()
  # Get root element
  production = DOMTree.documentElement
  # make <month> element
  monthElement = doc.createElement("month")
  # Set its attributes id to month
  month = sheet.getCellRangeByName(MONTH_CELLNO).String
  monthElement.setAttribute("id", month)
  # Read carton number
  currentRow = CARTON_NO_COL + str(CARTON_NO_ROW_START)
  cartonNo = sheet.getCellRangeByName(currentRow).String
  while cartonNo != "":
    # Create <carton> element and set its id
    cartonElement = doc.createElement("carton")
    cartonElement.setAttribute("id", cartonNo)
    # Create <modelno> element and set its data
    modelnoElement = doc.createElement("modelno")
    text = sheet.getCellRangeByName(MODEL_NO_COL + str(CARTON_NO_ROW_START)).String
    textNode = doc.createTextNode(text)
    modelnoElement.appendChild(textNode)
    # Create <model> element and set its data
    modelElement = doc.createElement("model")
    text = sheet.getCellRangeByName(MODEL_NAME_COL + str(CARTON_NO_ROW_START)).String
    textNode = doc.createTextNode(text)
    modelElement.appendChild(textNode)
    # Create <pv> element and set its data
    pvElement = doc.createElement("pv")
    text = sheet.getCellRangeByName(PV_COL + str(CARTON_NO_ROW_START)).String
    textNode = doc.createTextNode(text)
    pvElement.appendChild(textNode)
    # Create <price> element and set its data
    priceElement = doc.createElement("price")
    text = sheet.getCellRangeByName(PRICE_COL + str(CARTON_NO_ROW_START)).String
    textNode = doc.createTextNode(text)
    priceElement.appendChild(textNode)
    # Append element to <carton> element
    cartonElement.appendChild(modelnoElement)
    cartonElement.appendChild(modelElement)
    cartonElement.appendChild(pvElement)
    cartonElement.appendChild(priceElement)
    # Append <carton> element to <month> element
    monthElement.appendChild(cartonElement)
    # Read next carton row
    CARTON_NO_ROW_START += 1
    currentRow = CARTON_NO_COL + str(CARTON_NO_ROW_START)
    cartonNo = sheet.getCellRangeByName(currentRow).String
  # Append to root element
  production.appendChild(monthElement)
  # Write xml file
  file_object = open(filePath, "wt") # important observe 't' in 'wt' string
  DOMTree.writexml(file_object) #, indent=" ", addindent=" ", newl="\n")
  file_object.close()
  # Display info box notifing xml has updated
  infoBox("xml file updated", "Success")
  return None
  
# Function to update sqlite database file for monthly production cost
def updateMasterDBFile(*args):
  PRODUCTION_SHEET = 0
  # Initalize and get required context
  model = initalize()
  sheet = model.Sheets.getByIndex(PRODUCTION_SHEET);
  sheet.getCellRangeByName("A2").String = "I am update DB button"
  return None
  

