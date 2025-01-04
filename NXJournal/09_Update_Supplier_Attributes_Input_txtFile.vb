Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF
imports system.io
imports system.windows.forms
Imports System.Math
Imports Nxopen.Drawings
Imports NXOpen.Assemblies

Module NXJournal

Public newbasnum as string
Public bnum as string
Public sht as string
Public DBR as string

Sub Main (ByVal args() As String) 

Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
'Dim workPart As NXOpen.Part = theSession.Parts.Work
'
'Dim displayPart As NXOpen.Part = theSession.Parts.Display
Dim lw As ListingWindow = theSession.ListingWindow

Dim fpath As String

fpath = My.Computer.FileSystem.SpecialDirectories.Desktop & "\open_supplierattributes.txt"
'fpath = "T:\Gopinadh\Conversion\RFQ.165\open_dxf.txt"

Dim fpaths() As String = file.readalllines(fpath)

'New code to save parts 

for each f as string in fpaths

IF(F.LENGTH > 5)

Dim basePart2 As NXOpen.BasePart
Dim partLoadStatus2 As NXOpen.PartLoadStatus
basePart2 = theSession.Parts.OpenBaseDisplay(f.replace("""",""), partLoadStatus2)

'my code

'Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
Dim workPart As NXOpen.Part = theSession.Parts.Work

Dim displayPart As NXOpen.Part = theSession.Parts.Display



Dim markId1 As NXOpen.Session.UndoMarkId = Nothing
markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Start")

Dim objects1(0) As NXOpen.NXObject
objects1(0) = workPart
Dim attributePropertiesBuilder1 As NXOpen.AttributePropertiesBuilder = Nothing
attributePropertiesBuilder1 = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, objects1, NXOpen.AttributePropertiesBuilder.OperationType.None)

attributePropertiesBuilder1.IsArray = False

attributePropertiesBuilder1.IsArray = False

attributePropertiesBuilder1.IsArray = False

attributePropertiesBuilder1.DataType = NXOpen.AttributePropertiesBaseBuilder.DataTypeOptions.String

attributePropertiesBuilder1.Units = "MilliMeter"

Dim objects2(0) As NXOpen.NXObject
objects2(0) = workPart
Dim massPropertiesBuilder1 As NXOpen.MassPropertiesBuilder = Nothing
massPropertiesBuilder1 = workPart.PropertiesManager.CreateMassPropertiesBuilder(objects2)

Dim selectNXObjectList1 As NXOpen.SelectNXObjectList = Nothing
selectNXObjectList1 = massPropertiesBuilder1.SelectedObjects

Dim objects3() As NXOpen.NXObject
objects3 = selectNXObjectList1.GetArray()

massPropertiesBuilder1.LoadPartialComponents = True

massPropertiesBuilder1.Accuracy = 0.98999999999999999

Dim objects4(0) As NXOpen.NXObject
objects4(0) = workPart
Dim previewPropertiesBuilder1 As NXOpen.PreviewPropertiesBuilder = Nothing
previewPropertiesBuilder1 = workPart.PropertiesManager.CreatePreviewPropertiesBuilder(objects4)

Dim objects5(0) As NXOpen.NXObject
objects5(0) = workPart
attributePropertiesBuilder1.SetAttributeObjects(objects5)

attributePropertiesBuilder1.Units = "MilliMeter"

theSession.SetUndoMarkName(markId1, "Displayed Part Properties Dialog")

attributePropertiesBuilder1.DateValue.DateItem.Day = NXOpen.DateItemBuilder.DayOfMonth.Day12

attributePropertiesBuilder1.DateValue.DateItem.Month = NXOpen.DateItemBuilder.MonthOfYear.Sep

attributePropertiesBuilder1.DateValue.DateItem.Year = "2017"

attributePropertiesBuilder1.DateValue.DateItem.Time = "00:00:00"

massPropertiesBuilder1.UpdateOnSave = NXOpen.MassPropertiesBuilder.UpdateOptions.No

attributePropertiesBuilder1.Category = "Supplier Attributes"

attributePropertiesBuilder1.StringValue = "Sup_Name"

attributePropertiesBuilder1.Title = "Sup_Name"

attributePropertiesBuilder1.StringValue = "KK"

Dim markId2 As NXOpen.Session.UndoMarkId = Nothing
markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Displayed Part Properties")

Dim nXObject1 As NXOpen.NXObject = Nothing
nXObject1 = attributePropertiesBuilder1.Commit()

Dim updateoption1 As NXOpen.MassPropertiesBuilder.UpdateOptions = Nothing
updateoption1 = massPropertiesBuilder1.UpdateOnSave

Dim nXObject2 As NXOpen.NXObject = Nothing
nXObject2 = massPropertiesBuilder1.Commit()

workPart.PartPreviewMode = NXOpen.BasePart.PartPreview.None

Dim nXObject3 As NXOpen.NXObject = Nothing
nXObject3 = previewPropertiesBuilder1.Commit()

Dim id1 As NXOpen.Session.UndoMarkId = Nothing
id1 = theSession.GetNewestUndoMark(NXOpen.Session.MarkVisibility.Visible)

Dim nErrs1 As Integer = Nothing
nErrs1 = theSession.UpdateManager.DoUpdate(id1)

theSession.DeleteUndoMark(markId2, Nothing)

theSession.SetUndoMarkName(id1, "Displayed Part Properties")

attributePropertiesBuilder1.Destroy()

massPropertiesBuilder1.Destroy()

previewPropertiesBuilder1.Destroy()

Dim markId3 As NXOpen.Session.UndoMarkId = Nothing
markId3 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Start")

Dim objects6(0) As NXOpen.NXObject
objects6(0) = workPart
Dim attributePropertiesBuilder2 As NXOpen.AttributePropertiesBuilder = Nothing
attributePropertiesBuilder2 = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, objects6, NXOpen.AttributePropertiesBuilder.OperationType.None)

attributePropertiesBuilder2.IsArray = False

attributePropertiesBuilder2.IsArray = False

attributePropertiesBuilder2.IsArray = False

attributePropertiesBuilder2.DataType = NXOpen.AttributePropertiesBaseBuilder.DataTypeOptions.String

attributePropertiesBuilder2.Units = "MilliMeter"

Dim objects7(0) As NXOpen.NXObject
objects7(0) = workPart
Dim massPropertiesBuilder2 As NXOpen.MassPropertiesBuilder = Nothing
massPropertiesBuilder2 = workPart.PropertiesManager.CreateMassPropertiesBuilder(objects7)

Dim selectNXObjectList2 As NXOpen.SelectNXObjectList = Nothing
selectNXObjectList2 = massPropertiesBuilder2.SelectedObjects

Dim objects8() As NXOpen.NXObject
objects8 = selectNXObjectList2.GetArray()

massPropertiesBuilder2.LoadPartialComponents = True

massPropertiesBuilder2.Accuracy = 0.98999999999999999

Dim objects9(0) As NXOpen.NXObject
objects9(0) = workPart
Dim previewPropertiesBuilder2 As NXOpen.PreviewPropertiesBuilder = Nothing
previewPropertiesBuilder2 = workPart.PropertiesManager.CreatePreviewPropertiesBuilder(objects9)

Dim objects10(0) As NXOpen.NXObject
objects10(0) = workPart
attributePropertiesBuilder2.SetAttributeObjects(objects10)

attributePropertiesBuilder2.Units = "MilliMeter"

theSession.SetUndoMarkName(markId3, "Displayed Part Properties Dialog")

attributePropertiesBuilder2.DateValue.DateItem.Day = NXOpen.DateItemBuilder.DayOfMonth.Day12

attributePropertiesBuilder2.DateValue.DateItem.Month = NXOpen.DateItemBuilder.MonthOfYear.Sep

attributePropertiesBuilder2.DateValue.DateItem.Year = "2017"

attributePropertiesBuilder2.DateValue.DateItem.Time = "00:00:00"

massPropertiesBuilder2.UpdateOnSave = NXOpen.MassPropertiesBuilder.UpdateOptions.No

' ----------------------------------------------
'   Dialog Begin Displayed Part Properties
' ----------------------------------------------
attributePropertiesBuilder2.Category = "Supplier Attributes"

attributePropertiesBuilder2.Title = "Sup_Version"

Dim markId4 As NXOpen.Session.UndoMarkId = Nothing
markId4 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Displayed Part Properties")

attributePropertiesBuilder2.StringValue = "001"

theSession.DeleteUndoMark(markId4, Nothing)

Dim markId5 As NXOpen.Session.UndoMarkId = Nothing
markId5 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Displayed Part Properties")

Dim nXObject4 As NXOpen.NXObject = Nothing
nXObject4 = attributePropertiesBuilder2.Commit()

Dim updateoption2 As NXOpen.MassPropertiesBuilder.UpdateOptions = Nothing
updateoption2 = massPropertiesBuilder2.UpdateOnSave

Dim nXObject5 As NXOpen.NXObject = Nothing
nXObject5 = massPropertiesBuilder2.Commit()

workPart.PartPreviewMode = NXOpen.BasePart.PartPreview.None

Dim nXObject6 As NXOpen.NXObject = Nothing
nXObject6 = previewPropertiesBuilder2.Commit()

Dim id2 As NXOpen.Session.UndoMarkId = Nothing
id2 = theSession.GetNewestUndoMark(NXOpen.Session.MarkVisibility.Visible)

Dim nErrs2 As Integer = Nothing
nErrs2 = theSession.UpdateManager.DoUpdate(id2)

theSession.DeleteUndoMark(markId5, Nothing)

theSession.SetUndoMarkName(id2, "Displayed Part Properties")

attributePropertiesBuilder2.Destroy()

massPropertiesBuilder2.Destroy()

previewPropertiesBuilder2.Destroy()

theSession.CleanUpFacetedFacesAndEdges()




Dim partSaveStatus1 As NXOpen.PartSaveStatus = Nothing
partSaveStatus1 = workPart.Save(NXOpen.BasePart.SaveComponents.False, NXOpen.BasePart.CloseAfterSave.False)

theSession.Parts.CloseAll(NXOpen.BasePart.CloseModified.CloseModified, Nothing)


ELSE

'MSGBOX("CHECK FILE PATHS")

END IF


next

'close

'next

End Sub

Sub reportComponentChildren( ByVal comp As Component, ByVal indent As Integer) 
For Each child As Component In comp.GetChildren()            
newbasnum = New String(" ", indent * 2) & child.DisplayName()
next              
end sub


End Module
