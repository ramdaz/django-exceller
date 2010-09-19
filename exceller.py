
from django.db.models import Q


import os,datetime
from pyExcelerator import *
from django.db.models.query import QuerySet
from pyExcelerator.Style import *

class Exceller:
	def __init__(self, queryset, fields, filename, path=None):
		self.queryset =queryset
		self.fields=fields
		
		self.filename = str(filename)
		if not self.filename.endswith(".xls"):
			self.filename+=self.filename +".xls"
		
		self.path = str(path)
		
	def check_integrity(self):
		if isinstance(self.queryset, QuerySet):
			pass
		else:
			raise "Pass Proper QuerySet"
		
		
		if isinstance(self.fields, tuple) or isinstance(self.fields, list):
			pass
		else:
			raise "Pass Proper QuerySet"
		if isinstance(self.filename,str):
			pass
		else:
			raise "The path should be a valid String"
		if os.path.exists(self.path):
			pass
		else:
			raise "The path should be a valid path"
		return True
		
	
	def prepare_fields(self):
		if self.queryset.count()>0:
			self.field_set=[]
			for field in self.fields:
			
				
				if field in self.queryset.values().field_names:
					self.field_set.append(field)
				else:
					try:
						x =self.queryset[0].__getattribute__(field).__str__()
				
						self.field_set.append(field)
					except AttributeError:
						pass
		else:
			self.field_set=[]
			
	def create_excel(self):
		self.prepare_fields()
		if len(self.field_set)>0:
			wb = Workbook()
			ws= wb.add_sheet('Sheet1')
			rowcount =len(self.field_set)
			colcount = self.queryset.count()
			for field in self.field_set:
				ws.write(0,self.field_set.index(field), field)
			
			for i in range(0, self.queryset.count()):
				for field in self.field_set:
					ws.write(i+1, self.field_set.index(field), self.queryset[i].__getattribute__(field).__str__())
			wb.save(self.path+"/" +self.filename)
			return "The Excel is Ready"
		else:
			return "The Field Sets are not prepared"
	
			
				
		
		
			
		