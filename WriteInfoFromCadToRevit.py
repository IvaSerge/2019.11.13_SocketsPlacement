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

import math

def mm_to_ft(mm):
	return 3.2808*mm/1000

def ft_to_mm(ft):
	return ft*304.8

def SetpParVal(elem, name, pValue):
	global doc
	try:
		TransactionManager.Instance.EnsureInTransaction(doc)
		elem.LookupParameter(name).Set(pValue)
		TransactionManager.Instance.TransactionTaskDone()
	except: pass



class rvtElem:

	def __init__(self, excelInfo):
		#Excel elements
		self.fName = i[8]
		self.tName = i[9]
		self.lvl = i[10]
		self.h = i[11]
		self.x = mm_to_ft(float(i[4]))
		self.y = mm_to_ft(float(i[5]))
		self.rotation = i[6]
		
		#Revit elements
		global doc
		self.doc = doc
		self.rvtType = self._getType()
		self.rvtLvl = self._getLevel()
		self.rvtIns = self._newIns()
		
		self._setRotation()

	def _getType(self): 
		"""Найти в проекте семейство и тип"""
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

	def _getLevel(self): 
		"""Найти в проекте уровень"""
		testFam = BuiltInParameter.DATUM_TEXT
		pvpF = ParameterValueProvider(ElementId(int(testFam)))
		fnrvStr = FilterStringEquals()
		filterC = ElementParameterFilter(FilterStringRule(
							pvpF, fnrvStr, self.lvl, False))
		rvtLvl = FilteredElementCollector(self.doc)\
			.OfCategory(BuiltInCategory.OST_Levels)\
			.WherePasses(filterC)\
			.FirstElement()
		return rvtLvl

	def _newIns(self):
		pt = XYZ(self.x, self.y, 0)
		tp = self.rvtType
		if not (tp.IsActive):
			tp.Activate()
			doc.Regenerate()
		lvl = self.rvtLvl
		str = Structure.StructuralType.NonStructural
		elem = self.doc.Create.NewFamilyInstance(pt, tp, lvl, str)
		return elem
	
	def _setRotation(self):
		global doc
		pnt1 = self.rvtIns.Location.Point
		pnt2 = XYZ(pnt1.X, pnt1.Y, pnt1.Z + 10)
		axis = Line.CreateBound(pnt1, pnt2)
		ang = self.rotation * math.pi/180 -  math.pi/2
		ElementTransformUtils.RotateElement(doc, self.rvtIns.Id, axis, ang)


info = IN[0]
reload = IN[1]
info.pop(0)

TransactionManager.Instance.EnsureInTransaction(doc)

rvtObj = [rvtElem(i) for i in info]

#Вписать информацию в семейство
TransactionManager.Instance.EnsureInTransaction(doc)

OUT = map(lambda x:x.rotation, rvtObj)
