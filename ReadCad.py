import clr
#clr.AddReference('RevitAPI')
#import Autodesk
#from Autodesk.Revit.DB import *

clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReferenceToFileAndPath(
"C:\\Program Files\\Autodesk\\AutoCAD 2017\\Autodesk.AutoCAD.Interop")
from Autodesk.AutoCAD.Interop import *

import System
from System import *

import math
from math import pi

import re
from re import *

acadApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Autocad.Application")
acadDoc = acadApp.ActiveDocument
modSpace = acadDoc.ModelSpace

def getAttrValueByTagName(_blinfo):
	blRefList = _blinfo[0]
	atrName = _blinfo[1]
	attrList = blRefList.GetAttributes()
	attr = [i for i in attrList if i.TagString == atrName]
	if attr:
		attrVal = attr[0].TextString
	else:
		attrVal = None
	return attrVal

def getFamilyType(_instalType):
	if _instalType == "1 ":
		defOut = ["_1___",
				"44 IP 2x"]
	elif _instalType == "3 ":
		defOut = ["_3___",
				"3_16"]
	elif _instalType == "3 ":
		defOut = ["_3___","3"]
	elif _instalType == "1 ":
		defOut = ["_3___","1"]
	else:
		defOut = ["",""]

	return defOut

def getElevation(_hight):
	check = _hight
	pattern = re.compile(r"\d+")
	if pattern.search(check):
		return pattern.search(check).group()
	else:
		return ""

def getPower(_power):
	check = _power
	pattern = re.compile(r"\d+(?:\.|,)?\d*")
	res = pattern.findall(check)
	if res and len(res) == 1:
		return round(float(res[0].replace(",","."))*1000, 2)
	else:
		return ""

reload = IN[0]

#obj in selection set
#objList = acadDoc.PickfirstSelectionSet

#obj in model
objList = [i for i in modSpace]

# 
blRefList = [i for i in objList if i.ObjectName == "AcDbBlockReference"]

#--------------------------------------------------------
#  
#--------------------------------------------------------
attrNames = [".,"]*(len(blRefList))
blPower = map(getAttrValueByTagName, zip(blRefList, attrNames))

#    
filtered = [i for i in zip(blRefList, blPower) if i[1]]
blRefList = [i[0] for i in filtered]
blPower = [i[1] for i in filtered]

attrNames = [".__,"]*(len(blRefList))
blH = map(getAttrValueByTagName, zip(blRefList, attrNames))

attrNames = ["._"]*(len(blRefList))
blConnection = map(getAttrValueByTagName, zip(blRefList, attrNames))

attrNames = [""]*(len(blRefList))
blName = map(getAttrValueByTagName, zip(blRefList, attrNames))

#--------------------------------------------------------
#      
#--------------------------------------------------------
pntList = [i.InsertionPoint for i in blRefList]
listX = [i[0] for i in pntList]
listY = [i[1] for i in pntList]
rotationList = [round(i.Rotation/pi*180) for i in blRefList]

#--------------------------------------------------------
#  
#--------------------------------------------------------

#    
familyTypeList = map(getFamilyType,blConnection)
familyList = [i[0] for i in familyTypeList]
typeList = [i[1] for i in familyTypeList]

#__
rvtElevation = map(getElevation, blH)

#
rvtPower = map(getPower, blPower)

#--------------------------------------------------------
#    
#--------------------------------------------------------
blPower.insert(0,"")
blH.insert(0,"")
blConnection.insert(0," ")
blName.insert(0,"")
listX.insert(0," X")
listY.insert(0," Y")
rotationList.insert(0,"")
familyList.insert(0,"_")
typeList.insert(0,"_")
emptyColumn = [""] * len(blName)
rvtElevation.insert(0,"_")
rvtPower.insert(0,"_")

#--------------------------------------------------------
# ,    
#--------------------------------------------------------

rvtLevel = [""] * len(blName)
rvtLevel[0] = "_"

rvtCos = [""] * len(blName)
rvtCos[0] = "_"

rvtLoadClass = [""] * len(blName)
rvtLoadClass[0] = "_"

OUT = zip(blName, blPower, blH, blConnection, listX, listY, rotationList,
emptyColumn, familyList, typeList, rvtLevel, rvtElevation, rvtPower,
rvtCos, rvtLoadClass)