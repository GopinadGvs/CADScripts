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

fpath = My.Computer.FileSystem.SpecialDirectories.Desktop & "\open_pdf.txt"

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

REM Dim printPDFBuilder1 As NXOpen.PrintPDFBuilder
REM printPDFBuilder1 = workPart.PlotManager.CreatePrintPdfbuilder()

REM printPDFBuilder1.Scale = 1.0

REM printPDFBuilder1.Size = NXOpen.PrintPDFBuilder.SizeOption.ScaleFactor

REM printPDFBuilder1.OutputText = NXOpen.PrintPDFBuilder.OutputTextOption.Polylines

REM printPDFBuilder1.Units = NXOpen.PrintPDFBuilder.UnitsOption.English

REM printPDFBuilder1.XDimension = 8.5

REM printPDFBuilder1.YDimension = 11.0

REM printPDFBuilder1.RasterImages = True

REM theSession.SetUndoMarkName(markId1, "Export PDF Dialog")

REM Dim markId2 As NXOpen.Session.UndoMarkId
REM markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Export PDF")

REM theSession.DeleteUndoMark(markId2, Nothing)

REM Dim markId3 As NXOpen.Session.UndoMarkId
REM markId3 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Export PDF")

REM printPDFBuilder1.Watermark = ""

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
		
		Dim printPDFBuilder1 As NXOpen.PrintPDFBuilder
printPDFBuilder1 = workPart.PlotManager.CreatePrintPdfbuilder()

printPDFBuilder1.Scale = 1.0

printPDFBuilder1.Size = NXOpen.PrintPDFBuilder.SizeOption.ScaleFactor

printPDFBuilder1.OutputText = NXOpen.PrintPDFBuilder.OutputTextOption.Polylines

printPDFBuilder1.Units = NXOpen.PrintPDFBuilder.UnitsOption.English

printPDFBuilder1.XDimension = 8.5

printPDFBuilder1.YDimension = 11.0

printPDFBuilder1.RasterImages = True

theSession.SetUndoMarkName(markId1, "Export PDF Dialog")

Dim markId2 As NXOpen.Session.UndoMarkId
markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Export PDF")

theSession.DeleteUndoMark(markId2, Nothing)

Dim markId3 As NXOpen.Session.UndoMarkId
markId3 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Export PDF")

printPDFBuilder1.Watermark = ""

		'msgbox(sheet.name)

		REM i = i+1

		REM Next


printPDFBuilder1.SourceBuilder.SetSheets(sheets1)

dim filnm as string = IO.Path.GetFileName(displaypart.FullPath).Replace("_dwg.prt","")

dim udir as string = ""

if filnm.length > 15

udir = My.Computer.FileSystem.SpecialDirectories.Desktop & "\pdf\" & filnm.substring(14,3)

else

udir = My.Computer.FileSystem.SpecialDirectories.Desktop & "\pdf"

end if


Dim dir as string = udir

Dim fil as string = udir & "\" & IO.Path.GetFileName(displaypart.FullPath).Replace(".prt","").Replace("_DWG1","").Replace("_DWG","") & "_V001_S" & sheet.name.Substring(6,3)& ".pdf"

If Directory.Exists(dir)

else

Directory.CreateDirectory(dir)

end if

If File.Exists(fil)

File.Delete(fil)

MessageBox.Show("Replacing Existing Drawing File.","GopinadhGvs.",MessageBoxButtons.OK,MessageBoxIcon.Information)

'msgbox("Replacing Existing file")

else

end if



'printPDFBuilder1.Filename = My.Computer.FileSystem.SpecialDirectories.Desktop & "\pdf\" & IO.Path.GetFileName(displaypart.FullPath).Replace(".prt","") &".pdf"

printPDFBuilder1.Filename = fil


Dim nXObject1 As NXOpen.NXObject
nXObject1 = printPDFBuilder1.Commit()

i = i+1

printPDFBuilder1.Destroy()


Next

'theSession.DeleteUndoMark(markId3, Nothing)

'theSession.SetUndoMarkName(markId1, "Export PDF")

'printPDFBuilder1.Destroy()

'theSession.DeleteUndoMark(markId1, Nothing)

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
