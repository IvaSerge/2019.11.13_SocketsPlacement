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

def SetupParVal(elem, name, pValue):
	global doc
	try:
		elem.LookupParameter(name).Set(pValue)
	except:
		bip = GetBuiltInParam(name)
		elem.get_Parameter(bip).Set(pValue)
	return elem

def GetBuiltInParam(paramName):
	builtInParams = System.Enum.GetValues(BuiltInParameter)
	param = []
	
	for i in builtInParams:
		if i.ToString() == paramName:
			param.append(i)
			break
		else:
			continue
	return param[0]

class rvtElem:

	def __init__(self, excelInfo):
		#Excel elements
		self.description = i[0]
		self.fName = i[8]
		self.tName = i[9]
		self.lvl = i[10]
		self.h = mm_to_ft(i[11])
		self.x = mm_to_ft(float(i[4]))
		self.y = mm_to_ft(float(i[5]))
		self.rotation = i[6]
		self.P = UnitUtils.ConvertToInternalUnits(
					i[12],DisplayUnitType.DUT_WATTS)
		self.cos = float(i[13])
		
		#Revit elements
		global doc
		self.doc = doc
		self.rvtType = self._getType()
		self.rvtLvl = self._getLevel()
		self.Pclass = self._getElClass(i[14])
		self.rvtIns = None


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
		if not rvtLvl:
			raise ValueError("No correct level found")
		return rvtLvl

	def _getElClass(self, className):
		rvtClass = FilteredElementCollector(self.doc)\
			.OfCategory(BuiltInCategory.OST_ElectricalLoadClassifications)\
			.ToElements()
		rvtClass = [i for i in rvtClass if i.Name == className][0]
		return rvtClass
		
	def newIns(self):
		"""Создание нового объекта"""
		pt = XYZ(self.x, self.y, 0)
		tp = self.rvtType
		if not (tp.IsActive):
			tp.Activate()
			doc.Regenerate()
		lvl = self.rvtLvl
		if not(lvl):
			raise ValueError("No correct level found")
			return None
		str = Structure.StructuralType.NonStructural
		elem = self.doc.Create.NewFamilyInstance(pt, tp, lvl, str)
		doc.Regenerate()
		self.rvtIns = elem
		return elem

	def setRotation(self):
		"""Корректировка угла поворота"""
		global doc
		pnt1 = self.rvtIns.Location.Point
		pnt2 = XYZ(pnt1.X, pnt1.Y, pnt1.Z + 10)
		axis = Line.CreateBound(pnt1, pnt2)
		ang = self.rotation * math.pi/180 -  math.pi/2
		ElementTransformUtils.RotateElement(doc, self.rvtIns.Id, axis, ang)

	def setParameters(self):
		elem = self.rvtIns
		SetupParVal(elem, "INSTANCE_FREE_HOST_OFFSET_PARAM", self.h)
		SetupParVal(elem, "Р_уст", self.P)
		SetupParVal(elem, "cos_Phi", self.cos)
		elem.LookupParameter("Класс_Нагрузки").Set(self.Pclass.Id)
		elem.get_Parameter(
				BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS
				).Set(self.description)
		doc.Regenerate()

info = IN[0]
reload = IN[1]
info.pop(0)
errorList = list()

#Начало транзакции
TransactionManager.Instance.EnsureInTransaction(doc)

rvtObj = [rvtElem(i) for i in info]
map(lambda x:x.newIns(), rvtObj)
map(lambda x:x.setRotation, rvtObj)
map(lambda x:x.setParameters, rvtObj)

#Окончание транзакции
TransactionManager.Instance.EnsureInTransaction(doc)

OUT = zip(map(lambda x:x.rvtIns, rvtObj))
