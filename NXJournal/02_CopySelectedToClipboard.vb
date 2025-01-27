Option Strict Off
Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports NXOpen
Imports NXOpenUI

Module Module2

Sub Main()

Dim theSession As Session = Session.GetSession()
Dim theUI As UI = UI.GetUI

If IsNothing(theSession.Parts.BaseWork) Then
'active part required
Return
End If

Dim numsel As Integer = theUI.SelectionManager.GetNumSelectedObjects()
Dim theComps As New List(Of Assemblies.Component)

For i As Integer = 0 To numsel - 1
Dim selObj As TaggedObject = theUI.SelectionManager.GetSelectedTaggedObject(i)
If TypeOf (selObj) Is Assemblies.Component Then
theComps.Add(selObj)
End If
Next

If theComps.Count = 0 Then
'no components found among the preselected objects
Return
End If

Dim fileNames As New List(Of String)
For Each tempComp As Assemblies.Component In theComps

Dim compPath As String = tempComp.Prototype.OwningPart.FullPath
If Not fileNames.Contains(compPath) Then
fileNames.Add(compPath)
End If

Next

If fileNames.Count = 0 Then
Return
End If

Dim d As New DataObject(DataFormats.FileDrop, fileNames.ToArray)
Clipboard.SetDataObject(d, True)

dim fnames as string = ""
dim temp as string = ""
dim j as integer = 0

for each f as string in fileNames
temp = path.getfilename(fileNames(j))
fnames = fnames & Vblf & temp
j=j+1
next

MessageBox.Show(fileNames.Count & " Files Copied to Clipboard" & fnames, "Clipboard", _
                        MessageBoxButtons.OK, MessageBoxIcon.Information)

'MessageBox.Show(fileNames.Count & " Files Copied to Clipboard")

End Sub

Public Function GetUnloadOption(ByVal dummy As String) As Integer

'Unloads the image immediately after execution within NX
GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately

End Function

End Module