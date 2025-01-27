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

fpath = My.Computer.FileSystem.SpecialDirectories.Desktop & "\open_dxf.txt"

Dim fpaths() As String = file.readalllines(fpath)

for each f as string in fpaths

Dim basePart2 As NXOpen.BasePart
Dim partLoadStatus2 As NXOpen.PartLoadStatus
basePart2 = theSession.Parts.OpenBaseDisplay(f.replace("""",""), partLoadStatus2)

'my code

Dim workPart As NXOpen.Part = theSession.Parts.Work
'
Dim displayPart As NXOpen.Part = theSession.Parts.Display

Dim markId1 As NXOpen.Session.UndoMarkId
markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Start")

Dim sheet As NXOpen.Drawings.DrawingSheet

Dim J As Integer = 0
For Each sheet In workPart.DrawingSheets
J = J+1
Next

'sheets1.Sort(AddressOf CompareSheetNames)


		Dim mySheets As New List(Of Drawings.DrawingSheet)
 
        'add each sheet object to the list
        For Each tempSheet As Drawings.DrawingSheet In theSession.Parts.Work.DrawingSheets
            mySheets.Add(tempSheet)
        Next
 
        'list the sheet names in the "as-added" order
        'lw.WriteLine("Original sheet order")
        For Each tempSheet As Drawings.DrawingSheet In mySheets
            'lw.WriteLine(tempSheet.Name)
        Next
        'lw.WriteLine("")
 
        'sort list alphabetically by sheet name (according to our function)
        mySheets.Sort(AddressOf CompareSheetNames)
 
        'list the sorted sheet names
        'lw.WriteLine("Sorted Sheet Order")
        For Each tempSheet As Drawings.DrawingSheet In mySheets
            'lw.WriteLine(tempSheet.Name)
        Next
        'lw.WriteLine("")	
		
	
		'Dim sheets1(J-1) As NXOpen.NXObject
		Dim sheets1(0) As NXOpen.NXObject

		Dim i As Integer = 0

For Each sheet In mySheets

		sheets1(0) = sheet
		'msgbox(sheet.name)

Dim dxfdwgCreator1 As NXOpen.DxfdwgCreator = Nothing
dxfdwgCreator1 = theSession.DexManager.CreateDxfdwgCreator()

dxfdwgCreator1.ExportData = NXOpen.DxfdwgCreator.ExportDataOption.Drawing

dxfdwgCreator1.AutoCADRevision = NXOpen.DxfdwgCreator.AutoCADRevisionOptions.R2004

dxfdwgCreator1.ViewEditMode = True

dxfdwgCreator1.FlattenAssembly = True

dxfdwgCreator1.ExportScaleValue = "1:1"

dxfdwgCreator1.SettingsFile = "C:\Program Files\Siemens\NX 12.0\dxfdwg\dxfdwg.def"

dxfdwgCreator1.OutputFile = "G:\Cat Data_728\DXF\F58000102175801450001_DWG.dxf"

dxfdwgCreator1.ObjectTypes.Curves = True

dxfdwgCreator1.ObjectTypes.Annotations = True

dxfdwgCreator1.ObjectTypes.Structures = True

dxfdwgCreator1.FlattenAssembly = False

dxfdwgCreator1.InputFile = displaypart.FullPath

dxfdwgCreator1.OutputFile = "C:\Daimler\NX11_Build_16_2_1\nx_server\start_apps\F58000102142801240081_DWG.dxf"

Dim markId2 As NXOpen.Session.UndoMarkId = Nothing
markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "AutoCAD DXF/DWG Export Wizard ")

theSession.DeleteUndoMark(markId2, Nothing)

Dim markId3 As NXOpen.Session.UndoMarkId = Nothing
markId3 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "AutoCAD DXF/DWG Export Wizard ")

dxfdwgCreator1.WidthFactorMode = NXOpen.DxfdwgCreator.WidthfactorMethodOptions.AutomaticCalculation

dxfdwgCreator1.LayerMask = "1-256"

dxfdwgCreator1.DrawingList = sheet.name

dim filnm as string = IO.Path.GetFileName(displaypart.FullPath).Replace("dwg.prt","")

REM dxfdwgCreator1.OutputFile = My.Computer.FileSystem.SpecialDirectories.Desktop & "\DXF\" & filnm.substring(14,3) & "\" & IO.Path.GetFileName(displaypart.FullPath).Replace(".prt","") & ".dxf"

dim outloc as string = My.Computer.FileSystem.SpecialDirectories.Desktop & "\DXF\" & IO.Path.GetFileName(displaypart.FullPath).Replace(".prt","") & "_" & sheet.name & ".dxf"

dxfdwgCreator1.OutputFile = My.Computer.FileSystem.SpecialDirectories.Desktop & "\DXF\" & IO.Path.GetFileName(displaypart.FullPath).Replace(".prt","") & "_" & sheet.name & ".dxf"

Dim nXObject1 As NXOpen.NXObject = Nothing
nXObject1 = dxfdwgCreator1.Commit()

theSession.DeleteUndoMark(markId3, Nothing)

i = i+1

Do Until System.IO.File.Exists(outloc)

Loop

dxfdwgCreator1.Destroy()


Next

theSession.CleanUpFacetedFacesAndEdges()

theSession.Parts.CloseAll(NXOpen.BasePart.CloseModified.CloseModified, Nothing)

'

next

End Sub

Sub reportComponentChildren( ByVal comp As Component, ByVal indent As Integer) 
For Each child As Component In comp.GetChildren()            
newbasnum = New String(" ", indent * 2) & child.DisplayName()
next			
end sub

Private Function CompareSheetNames(ByVal x As Drawings.DrawingSheet, ByVal y As Drawings.DrawingSheet) As Integer
 
        'case-insensitive sort
        Dim myStringComp As StringComparer = StringComparer.CurrentCultureIgnoreCase
 
        'for a case-sensitive sort (A-Z then a-z), change the above option to:
        'Dim myStringComp As StringComparer = StringComparer.CurrentCulture
 
        Return myStringComp.Compare(x.Name, y.Name)
 
    End Function


End Module
