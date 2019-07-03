'''
This script will print out a list of selected revisions.

This script was created by Christopher Maltez circa June 16th, 2019
'''


#-------------------------------------------------------------------------------------
#this is the init.py script

import clr
import rpw
from rpw import revit, db, ui, DB, UI

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReferenceByPartialName('PresentationCore')
clr.AddReferenceByPartialName("PresentationFramework")
clr.AddReferenceByPartialName('System')
clr.AddReferenceByPartialName('System.Windows.Forms')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from Autodesk.Revit.DB.Analysis import *
from Autodesk.Revit.UI import *


from Autodesk.Revit import DB
from Autodesk.Revit import UI
from Autodesk.Revit import *


uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document


from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI import UIApplication


def alert(msg):
    TaskDialog.Show('RevitPythonShell', msg)


# def quit():
#     __window__.Close()
# exit = quit


def get_selected_elements(doc):
    """API change in Revit 2016 makes old method throw an error"""
    try:
        # Revit 2016
        return [doc.GetElement(id)
                for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
    except:
        # old method
        return list(__revit__.ActiveUIDocument.Selection.Elements)
selection = get_selected_elements(doc)
# convenience variable for first element in selection
if len(selection):
    s0 = selection[0]

#------------------------------------------------------------------------------

#this is the init.py script
import clr
from Autodesk.Revit.DB import ElementSet, ElementId

class RevitLookup(object):
    def __init__(self, uiApplication):
        '''
        for RevitSnoop to function properly, it needs to be instantiated
        with a reference to the Revit Application object.
        '''
        # find the RevitLookup plugin
        try:
			rlapp = [app for app in uiApplication.LoadedApplications
					 if app.GetType().Namespace == 'RevitLookup'
					 and app.GetType().Name == 'App'][0]
        except IndexError:
            self.RevitLookup = None
            return
        # tell IronPython about the assembly of the RevitLookup plugin
        clr.AddReference(rlapp.GetType().Assembly)
        import RevitLookup
        self.RevitLookup = RevitLookup
        # See note in CollectorExt.cs in the RevitLookup source:
        self.RevitLookup.Snoop.CollectorExts.CollectorExt.m_app = uiApplication
        self.revit = uiApplication

    def lookup(self, element):
        if not self.RevitLookup:
			print 'RevitLookup not installed. Visit https://github.com/jeremytammik/RevitLookup to install.'
			return
        if isinstance(element, int):
            element = self.revit.ActiveUIDocument.Document.GetElement(ElementId(element))
        if isinstance(element, ElementId):
            element = self.revit.ActiveUIDocument.Document.GetElement(element)
        if isinstance(element, list):
            elementSet = ElementSet()
            for e in element:
                elementSet.Insert(e)
            element = elementSet
        form = self.RevitLookup.Snoop.Forms.Objects(element)
        form.ShowDialog()
_revitlookup = RevitLookup(__revit__)
def lookup(element):
    _revitlookup.lookup(element)

#------------------------------------------------------------------------------

# a fix for the __window__.Close() bug introduced with the non-modal console
'''
class WindowWrapper(object):
    def __init__(self, win):
        self.win = win

    def Close(self):
        self.win.Dispatcher.Invoke(lambda *_: self.win.Close())

    def __getattr__(self, name):
        return getattr(self.win, name)
__window__ = WindowWrapper(__window__)
'''


#-------------------------------------------------------------------------------------

'''
This script will print out a list of selected revisions.

This script was created by Christopher Maltez circa June 16th, 2019
'''

import os
import csv
import rpw
from rpw.ui.forms import Console
from rpw.ui.forms import SelectFromList

import clr
from rpw import revit, db, ui, DB, UI

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReferenceByPartialName('PresentationCore')
clr.AddReferenceByPartialName("PresentationFramework")
clr.AddReferenceByPartialName('System')
clr.AddReferenceByPartialName('System.Windows.Forms')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from Autodesk.Revit.DB.Analysis import *
from Autodesk.Revit.UI import *

from Autodesk.Revit import DB
from Autodesk.Revit import UI



desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
filepath = os.path.join(desktop, 'RevisionClouds.csv')

#sl = FilteredElementCollector(doc)
#sl.OfCategory(BuiltInCategory.OST_Sheets)
#sheets = sl.OfCategory(BuiltInCategory.OST_Sheets)
#for x in sheets:
	#print(x.Id)
#for x in cl:
#	list(x.GetSheetIds())

#this below might be a lead
#clouds = db.Collector(of_category='Revision Clouds', is_type=False)
#for x in clouds:
	#print(x.OwnerViewId)
	#doc.GetElement(x.OwnerViewId).Title


cl = FilteredElementCollector(doc)
cl.OfCategory(BuiltInCategory.OST_RevisionClouds)
cl.WhereElementIsNotElementType()


def get_revCloudID(element):
	return element.Id

def get_revision(element, parameterName):
	return element.LookupParameter(parameterName).AsValueString()
	#need to check for zeros

def get_comments(element, parameterName):
	CommentString = element.LookupParameter("Comments").AsString()
	return CommentString
	#need to check for zeros
	
def get_comment_parent_sheet_name():
		try: 
			return doc.GetElement(list(element.GetSheetIds())[0]).Title
		except IndexError:
			return doc.GetElement(element.OwnerViewId).Title
	#need to check for zeros

def get_comment_parent_view_name():
	return doc.GetElement(element.OwnerViewId).Name
	#need to check for zeros
	
def get_comment_parent_view_number(element, parameterName):
	return doc.GetElement(element.OwnerViewId).LookupParameter("Detail Number").AsString()
	#need to check for zeros

revList = Revision.GetAllRevisionIds(doc)
#need to check for zeros


revNameList = []


for x in revList: 
	revName = doc.GetElement(x).Name
	revNameList.append(revName)
	#need to check for zeros
	

selectedRevisionValue = SelectFromList('Test Window', revNameList)
	     	
#with open(r'C:\Users\cmaltez\Desktop\PyRy\mycsv2.csv', 'wb') as f:
with open(filepath, 'wb') as f:
	fieldnames = [
		'RevCloudID',
		'Comment',
		'Revsion', 
		'Sheet Name', 
		'View Number', 
		'View Name'
		]
	thewriter = csv.DictWriter(f, fieldnames=fieldnames)
	thewriter.writeheader()


	for element in cl: 
		#need to check for zeros	
		if get_revision(element, "Revision") == selectedRevisionValue:
		#	if any(field.strip() for field in cl):
			thewriter.writerow({
				 'RevCloudID': get_revCloudID(element),
		         'Comment': get_comments(element, "Comments"), 
		         'Revsion': get_revision(element, "Revision"), 
		         'Sheet Name': get_comment_parent_sheet_name(), 
		         'View Number': get_comment_parent_view_number(element, "Detail Number"), 
		         'View Name': get_comment_parent_view_name()
		         })
		         
os.system("start EXCEL.EXE " + filepath)
	
'''
	for element in cl:
		if get_revision(element, "Revision") == selectedRevisionValue:
		
			print "Comment: " + get_comments(element, "Comments")
			print "Revision: " + get_revision(element, "Revision")
			print "Sheet Name: " + get_comment_parent_sheet_name()
			print "View Number: " + get_comment_parent_view_number(element, "Detail Number")
			print "View Name: " + get_comment_parent_view_name()
			print "\n"
	'''
	
	#else: 
	#	print("There are no revision clouds on the '" + selectedRevisionValue + "' revision.")


#TODO: CHECK FOR EMPTY REVISIONS
#TODO: EXPORT TO SCHEDULE TO KEEP IN REVIT
'''
#with open('C:\Users\cmaltez\Desktop\PyRy\mycsv.csv', 'w', newline='') as f:
with open(r'C:\Users\cmaltez\Desktop\PyRy\mycsv4.csv', 'w') as f:
	fieldnames = ['column1', 'column2', 'column3']
	thewriter = csv.DictWriter(f, fieldnames=fieldnames)
	thewriter.writeheader()
	for i in range(1, 10):
		thewriter.writerow({'column1':'one', 'column2':'two', 'column3':'three'})
'''

#TODO: EXPORT TO WORD DOC IN RMW FORMAT

#open function needs full filepath!