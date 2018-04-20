#Structure of xml
#Path to narration element: <ENVELOPE>/<BODY>/<IMPORTDATA>/<REQUESTDATA>/<TALLYMESSAGE>/<NARRATION>
#Path to voucher no: <ENVELOPE>/<BODY>/<IMPORTDATA>/<REQUESTDATA>/<TALLYMESSAGE>/<VOUCHERNUMBER>
#Path to outside labour charges: <ENVELOPE>/<BODY>/<IMPORTDATA>/<REQUESTDATA>/<TALLYMESSAGE>/<LEDGERENTRIES.LIST>/<AMOUNT> (get absolute value)
#Path to number of cycles: <ENVELOPE>/<BODY>/<IMPORTDATA>/<REQUESTDATA>/<TALLYMESSAGE>/<INVENTORYENTRIESIN.LIST>/<ACTUALQTY> (format: 24.00 Pcs)
#Path to amount: <ENVELOPE>/<BODY>/<IMPORTDATA>/<REQUESTDATA>/<TALLYMESSAGE><INVENTORYENTRIESIN.LIST>/<AMOUNT> (get absolute value)

import sys
from xml.dom.minidom import parse
import xml.dom.minidom
from com.sun.star.beans import PropertyValue
#shortcut used in creating file dialogs
createUnoService = (
        XSCRIPTCONTEXT
        .getComponentContext()
        .getServiceManager()
        .createInstance 
                    )

#function to load production data from XML exported by Tally
def Initalize():
  #get the doc from the scripting context which is made available to all scripts
  desktop = XSCRIPTCONTEXT.getDesktop()
  #desktop.createInstance("com.sun.star.ui.dialogs.FilePicker")
  model = desktop.getCurrentComponent()
  return model

def loadProductionData(*args):
  
  
  model = Initalize()
  sheet = model.Sheets.getByIndex(0);
  #Not happy with below 2 statement just a cheap quick fix to make things working in a hurry
  #TODO need to find a better solution to it  
  #[7:] syntax used to get rid of 'file://' part of file url returned by FilePicker
  filePath = FilePicker()[7:]
  #Replace a single space in folder name Kingston Accounts which is returned as %20 by FilePicker
  filePath = filePath.replace('%20', ' ')
  # Open XML document using minidom parser
  DOMTree = xml.dom.minidom.parse(filePath)
  collection = DOMTree.documentElement
  
  # Get list of all productions
  productions = collection.getElementsByTagName("TALLYMESSAGE")
  for production in productions:
    narration = production.getElementsByTagName("NARRATION")
    
    #check if narration tag is present
    if narration:
      # Production Voucher number
      pvNumber = production.getElementsByTagName("VOUCHERNUMBER")[0].childNodes[0].data
      #get cycle out tag
      cycle = production.getElementsByTagName("INVENTORYENTRIESIN.LIST")[0]
      #get amount
      cycle_amount_tag = cycle.getElementsByTagName("AMOUNT")[0]
      cycle_amount = float(cycle_amount_tag.childNodes[0].data)
    
      #get number of cycles
      cycle_num_tag = cycle.getElementsByTagName("ACTUALQTY")[0]
      cycle_num = cycle_num_tag.childNodes[0].data
      cycle_num_str = cycle_num_tag.childNodes[0].data
      cycle_num_str = cycle_num_str.strip()
      cycle_num = float(cycle_num_str.split(" ")[0])
    
      #get outside labour charges
      outside_labour_tag = production.getElementsByTagName("LEDGERENTRIES.LIST")[0]
      outside_labour_charges_tag = outside_labour_tag.getElementsByTagName("AMOUNT")[0]
      outside_labour_charges_val = float(outside_labour_charges_tag.childNodes[0].data)
      total_cost = abs((cycle_amount+outside_labour_charges_val)/cycle_num)
      total_cost = round(total_cost, 2)
      
      #write cycle cost to appropriate carton numbers
      #writeProductionValue(sheet, narration[0].childNodes[0].data, total_cost, pvNumber)
      writeProductionValue(narration[0].childNodes[0].data, total_cost, pvNumber)
  
  return None

def writeProductionValue(cartonDataStr, cycleCost, pvNumber):
  cartonNoStartRow = 3
  cartonCostCol = 'P'
  cartonPVCol = 'O'
  PRODUCTION_SHEET = 0
  # Initalize and get required context
  model = Initalize()
  sheet = model.Sheets.getByIndex(PRODUCTION_SHEET);
  # Opening bracket index
  openingBracketIndex = cartonDataStr.index("[")
  # Closing bracket index
  closingBracketIndex = cartonDataStr.index("]")
  cartonDataStr = cartonDataStr[openingBracketIndex+1:closingBracketIndex]
  cartonNumPairList = cartonDataStr.split(",")
  for cartonNumPair in cartonNumPairList:
    cartonNums = cartonNumPair.split(":")
    startCartonNo = int(cartonNums[0])
    endCartonNo = 0
    if len(cartonNums) > 1:
      endCartonNo = int(cartonNums[1])
    # If only one carton is present
    if endCartonNo == 0:
      # Write cycle cost
      cartonNo = startCartonNo + cartonNoStartRow
      # Wrtie cost
      cellNo = cartonCostCol + str(cartonNo)
      sheet.getCellRangeByName(cellNo).String = cycleCost
      # Wrtie pv number
      cellNo = cartonPVCol + str(cartonNo)
      sheet.getCellRangeByName(cellNo).String = pvNumber
    else: # there are list of cartons
      numOfCartons = endCartonNo - startCartonNo + 1
      for i in range(0, numOfCartons):
        cartonNo = cartonNoStartRow + startCartonNo + i
        # Write cost
        cellNo = cartonCostCol + str(cartonNo)
        sheet.getCellRangeByName(cellNo).String = cycleCost
        # Write pv number
        cellNo = cartonPVCol + str(cartonNo)
        sheet.getCellRangeByName(cellNo).String = pvNumber

def FilePicker(path=None, mode=1):

  filepicker = createUnoService("com.sun.star.ui.dialogs.FilePicker")
  filepicker.initialize( ( 0,) )
  filepicker.setMultiSelectionMode(False)
  filepicker.appendFilter("XML files (*.xml)", "*.xml")
  filepicker.setCurrentFilter("XML files (*.xml)")
  filepicker.execute()
  return filepicker.getFiles()[0]
  
