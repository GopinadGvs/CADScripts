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

fpath = My.Computer.FileSystem.SpecialDirectories.Desktop & "\open_jt.txt"

Dim fpaths() As String = file.readalllines(fpath)

for each f as string in fpaths

Dim basePart2 As NXOpen.BasePart
Dim partLoadStatus2 As NXOpen.PartLoadStatus
basePart2 = theSession.Parts.OpenBaseDisplay(f.replace("""",""), partLoadStatus2)

Dim workPart As NXOpen.Part = theSession.Parts.Work
'
Dim displayPart As NXOpen.Part = theSession.Parts.Display

'my code

Dim outpath = My.Computer.FileSystem.SpecialDirectories.Desktop & "\jt\" & IO.Path.GetFileName(displaypart.FullPath).Replace(".prt","") & ".jt"

If Directory.Exists(My.Computer.FileSystem.SpecialDirectories.Desktop & "\jt\")

else

Directory.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.Desktop & "\jt\")

end if


' ----------------------------------------------
'   Menu: File->Export->JT...
' ----------------------------------------------
Dim markId1 As NXOpen.Session.UndoMarkId
markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Start")

Dim jtCreator1 As NXOpen.JtCreator
jtCreator1 = theSession.PvtransManager.CreateJtCreator()

jtCreator1.IncludePmi = NXOpen.JtCreator.PmiOption.PartAndAsm

jtCreator1.TessOption = NXOpen.JtCreator.TessellationOption.Defined

jtCreator1.LighweightLabel = "JT_Facet"

jtCreator1.ConfigFile = "C:\Program Files\Siemens\NX 10.0\pvtrans\tessUG.config"

jtCreator1.JtfileStructure = NXOpen.JtCreator.FileStructure.Monolithic

jtCreator1.AutolowLod = True

jtCreator1.TessOption = NXOpen.JtCreator.TessellationOption.Nx

jtCreator1.PreciseGeom = True

theSession.SetUndoMarkName(markId1, "Export JT Dialog")

Dim listCreator1 As NXOpen.ListCreator
listCreator1 = jtCreator1.NewLevel()

listCreator1.TessOption = NXOpen.ListCreator.TessellationOption.Defined

listCreator1.Chordal = 0.001

listCreator1.Angular = 20.0

jtCreator1.LodList.Append(listCreator1)

Dim listCreator2 As NXOpen.ListCreator
listCreator2 = jtCreator1.NewLevel()

listCreator2.TessOption = NXOpen.ListCreator.TessellationOption.Defined

listCreator2.Chordal = 0.001

listCreator2.Angular = 20.0

jtCreator1.LodList.Append(listCreator2)

Dim listCreator3 As NXOpen.ListCreator
listCreator3 = jtCreator1.NewLevel()

listCreator3.TessOption = NXOpen.ListCreator.TessellationOption.Defined

listCreator3.Chordal = 0.001

listCreator3.Angular = 20.0

jtCreator1.LodList.Append(listCreator3)

listCreator2.Chordal = 0.0035

listCreator2.Angular = 0.0

listCreator2.Simplify = 0.4

listCreator2.AdvCompression = 0.5

listCreator3.Chordal = 0.01

listCreator3.Angular = 0.0

listCreator3.Simplify = 0.1

listCreator3.AdvCompression = 1.0

Dim markId2 As NXOpen.Session.UndoMarkId
markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Export JT")

theSession.DeleteUndoMark(markId2, Nothing)

Dim markId3 As NXOpen.Session.UndoMarkId
markId3 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Export JT")

jtCreator1.OutputJtFile = outpath

Dim nXObject1 As NXOpen.NXObject
nXObject1 = jtCreator1.Commit()

theSession.DeleteUndoMark(markId3, Nothing)

theSession.SetUndoMarkName(markId1, "Export JT")

jtCreator1.Destroy()

theSession.CleanUpFacetedFacesAndEdges()


'
theSession.Parts.CloseAll(NXOpen.BasePart.CloseModified.CloseModified, Nothing)

next

End Sub

Sub reportComponentChildren( ByVal comp As Component, ByVal indent As Integer) 
For Each child As Component In comp.GetChildren()            
newbasnum = New String(" ", indent * 2) & child.DisplayName()
next			
end sub


End Module
