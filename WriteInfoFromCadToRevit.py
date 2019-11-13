import clr
clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

import sys
pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import System
from System import Array

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

class rvtElem:

	def __init__(self, excelInfo):
		global doc
		self.doc = doc
		self.fName = i[8]
		self.tName = i[9]
		self.lvl = i[10]
		self.h = i[11]
		self.x = i[4]
		self.y = i[5]
		self.rotation = i[6]

	def getType(self): 
		testFam = BuiltInParameter.ALL_MODEL_FAMILY_NAME
		pvpF = ParameterValueProvider(ElementId(int(testFam)))
		fnrvStr = FilterStringEquals()
		filterF = ElementParameterFilter(FilterStringRule(
							pvpF, fnrvStr, self.fName, False))
		
		testType = BuiltInParameter.SYMBOL_NAME_PARAM
		pvpT = ParameterValueProvider(ElementId(int(testType)))
		filterT = ElementParameterFilter(FilterStringRule(
							pvpT, fnrvStr, self.tName, False))
		
		filterAnd = LogicalAndFilter(filterF, filterT)
		
		rvtType = FilteredElementCollector(self.doc)\
			.OfClass(FamilySymbol)\
			.WherePasses(filterAnd)\
			.FirstElement()
		
		return rvtType

info = IN[0]
reload = IN[1]

info.pop(0)

rvtObj = [rvtElem(i) for i in info]

#Найти в проекте семейство и тип
typList = map(lambda x:x.getType(), rvtObj)

#Преобразовать координаты и вставить по ним семейство

#Развернуть семейство

#Вписать информацию в семейство

OUT = typList
