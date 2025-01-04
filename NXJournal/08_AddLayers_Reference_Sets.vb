' NX 12.0.0.27
' Journal to Add to refset & to move Layer of RoughPart, Flamecut, Intersection Curves & Points to respective Layers
' Created by GopinadhGvs on Mon Dec 23 2018
 
Option Strict Off
 
Imports System
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Assemblies
Imports System.Collections.Generic
Imports NXOpenUI
 
Module NXJournal
 
Public BLayer as integer = 110
Public PLayer as integer = 41
Public ILayer as integer = 51

Public outobjects1(0) As NXOpen.DisplayableObject
Public pobjects1 As New List(Of NXOpen.DisplayableObject)
Public intobjects1 As New List(Of NXOpen.DisplayableObject)
Public bobjects1 As New List(Of NXOpen.DisplayableObject)
Public refSetObjs(0) As DisplayableObject

Public theSession As Session = Session.GetSession()
Public ufs As UFSession = UFSession.GetUFSession()
Public lw As ListingWindow = theSession.ListingWindow
Public theUfSession As UFSession = UFSession.GetUFSession

Public Complist As New List(Of Component)
Public Complistname As New List(Of string)
Public Complistunique As New List(Of string)
Public workPart As Part = theSession.Parts.Work
Public displayPart As Part = theSession.Parts.Display

Public mainassy as string
Public mirpart as integer
Public bool as integer
 
Sub Main()
REM 'Main try
Try
	
mainassy = IO.Path.GetFileNameWithoutExtension(workPart.FullPath)

Dim componentOrder1 As NXOpen.Assemblies.ComponentOrder = CType(workPart.ComponentAssembly.OrdersSet.FindObject("Alphanumeric"), NXOpen.Assemblies.ComponentOrder)
componentOrder1.Activate()

Dim c As ComponentAssembly = workPart.ComponentAssembly

if not IsNothing(c.RootComponent) then
ReportComponentChildren(c.RootComponent, 0)
else
'lw.WriteLine("Part has no components")
end if	

For Each compname As string In Complistname	
if (Complistunique.contains(compname))
else
Complistunique.add(compname)
end if		
Next

For Each nme As string In Complistunique
Dim part1 As NXOpen.Part = CType(theSession.Parts.FindObject(nme), NXOpen.Part)

Dim partLoadStatus1 As NXOpen.PartLoadStatus
Dim status1 As NXOpen.PartCollection.SdpsStatus
status1 = theSession.Parts.SetDisplay(part1, False, True, partLoadStatus1)

workPart = theSession.Parts.Work
displayPart = theSession.Parts.Display

'Try to check mirror part
Try
Dim Mirror As String
Mirror = workPart.GetStringAttribute("CSYSMirrorOption")
mirpart = 1
Catch ex As NXException
mirpart = 0
REM msgbox("Mirror Check: " & ex.message)
End Try
'Try to check mirror part
	
if (mirpart = 0)

REM pobjects1 = Nothing
REM intobjects1 = Nothing

'Try remove interface objects
REM try

REM Dim markId2 As NXOpen.Session.UndoMarkId = Nothing
REM markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Orphan Product Interface Remover")

REM Dim objectBuilder1 As NXOpen.Assemblies.ProductInterface.ObjectBuilder = Nothing
REM objectBuilder1 = workPart.ProductInterface.CreateObjectBuilderWithVersion(2)
REM Dim prodIntObjs1() As NXOpen.Assemblies.ProductInterface.InterfaceObject
REM prodIntObjs1 = objectBuilder1.QueryProductInterfaceObjects(workPart)
REM dim j as integer = 0
REM for each myobj As NXOpen.Assemblies.ProductInterface.InterfaceObject In prodIntObjs1
REM myobj = prodIntObjs1(j)
REM myobj.GetProductInterfaceObjectStatus
REM If myobj.GetProductInterfaceObjectStatus = "Orphan"
REM myobj.RemoveProductInterfaceObject()

REM Dim objects1(-1) As NXOpen.TaggedObject
REM Dim nErrs1 As Integer = Nothing
REM nErrs1 = theSession.UpdateManager.AddObjectsToDeleteList(objects1)

REM Dim notifyOnDelete2 As Boolean = Nothing
REM notifyOnDelete2 = theSession.Preferences.Modeling.NotifyOnDelete

REM Dim nErrs2 As Integer = Nothing
REM nErrs2 = theSession.UpdateManager.DoUpdate(markId2)

REM End If
REM j=j+1
REM Next
REM Catch ex As NXException
REM REM msgbox("Try remove interface: " & ex.message)
REM End Try
'Try remove interface objects


Dim displayModification1 As NXOpen.DisplayModification
displayModification1 = theSession.DisplayManager.NewDisplayModification()
displayModification1.ApplyToAllFaces = True
displayModification1.ApplyToOwningParts = True

'Get First Feature of Roughpart
Dim prevFeat As New list (of Features.Feature)
For Each tempFeat As Features.Feature In workPart.Features
If TypeOf (tempFeat) Is Features.FeatureGroup Then
Dim numFeatures As Integer
Dim setFeatureTags() As Tag
theUfSession.Modl.AskAllMembersOfSet(tempFeat.Tag, setFeatureTags, numFeatures)
Dim setFeatures As New List(Of Features.Feature)
if (tempFeat.name = "_Rohteil" or tempFeat.name = "_Rough_Part")
'get features from tags
For Each tempTag As Tag In setFeatureTags
	setFeatures.Add(Utilities.NXObjectManager.Get(tempTag))
Next
For Each setFeat As Features.Feature In setFeatures
If TypeOf (setFeat) Is Features.BodyFeature Then
prevFeat.Add(setFeat)
end if
Next
else
'msgbox("NoObjects found.")				
end if
End If 
Next

'Try to check body feature
try
if (prevFeat.count>0) then
dim fi as string = prevFeat(0).Getfeaturename.ToUpper()
Dim myFeatBody As features.Bodyfeature
myFeatBody = CType(prevFeat(0), Nxopen.Features.Bodyfeature)
Dim nxObj() As NXObject = myFeatBody.GetBodies()
For Each tempbody As Body In nxObj
Dim pobjects3(0) As NXOpen.DisplayableObject
pobjects3(0) = tempbody
REM bobjects1.Add(pobjects3(0))
refSetObjs(0) = pobjects3(0)
Next
end if
Catch ex As NXException
REM msgbox("Body feature check: " & ex.message)
End Try
'Try to check body feature


'Get First Feature of Roughpart

'Try Add Brennschablone
Try
'Exp & Out Brennschablone to 110 Layer
For Each bdy As NXOpen.Body In workPart.Bodies
dim nm as string = bdy.Name
if nm = "EXP_BRENNSCHABLONE"
bool=1
Dim expobjects1(0) As NXOpen.DisplayableObject
expobjects1(0) = bdy
displayModification1.NewLayer = BLayer
displayModification1.Apply(expobjects1)
REM else if nm = "OUT_BRENNSCHABLONE"
REM bool=1
REM outobjects1(0) = bdy
REM displayModification1.NewLayer = BLayer
REM displayModification1.Apply(outobjects1)
else
bool=0
end if
next
'Exp & Out Brennschablone to 110 Layer
Catch ex As NXException
REM msgbox("Add Brennschablone: " & ex.message)
End Try
'Try Add Brennschablone


'Get Body of output linked geometry
Dim prevFeat1 As New list (of Features.Feature)
For Each tempFeat As Features.Feature In workPart.Features
If TypeOf (tempFeat) Is Features.FeatureGroup Then
Dim numFeatures As Integer
Dim setFeatureTags() As Tag
theUfSession.Modl.AskAllMembersOfSet(tempFeat.Tag, setFeatureTags, numFeatures)
Dim setFeatures As New List(Of Features.Feature)
if (tempFeat.name = "_Output_verlinkte_Geometrie" or tempFeat.name = "_Output_linked_Geometry")
'get features from tags
For Each tempTag As Tag In setFeatureTags
	setFeatures.Add(Utilities.NXObjectManager.Get(tempTag))
Next
For Each setFeat As Features.Feature In setFeatures
If TypeOf (setFeat) Is Features.BodyFeature Then
prevFeat1.Add(setFeat)
end if
Next
else
'msgbox("NoObjects found.")				
end if
End If 
Next

'Try to check body feature output linked geometry
try
if (prevFeat1.count>0) then
dim fi as string = prevFeat1(0).Getfeaturename.ToUpper()
Dim myFeatBody As features.Bodyfeature
myFeatBody = CType(prevFeat1(0), Nxopen.Features.Bodyfeature)
Dim nxObj() As NXObject = myFeatBody.GetBodies()
For Each tempbody As Body In nxObj
REM Dim pobjects3(0) As NXOpen.DisplayableObject
REM pobjects3(0) = tempbody
REM REM bobjects1.Add(pobjects3(0))
REM refSetObjs(0) = pobjects3(0)
outobjects1(0) = tempbody
displayModification1.NewLayer = BLayer
displayModification1.Apply(outobjects1)
Next
end if
Catch ex As NXException
REM msgbox("Body feature check output linked geometry: " & ex.message)
End Try
'Try to check body feature output linked geometry

'Points & Intersection Curves to 41 & 51 Layer
For Each tempFeat As Features.Feature In workPart.Features
If TypeOf (tempFeat) Is Features.FeatureGroup Then
Dim numFeatures As Integer
Dim setFeatureTags() As Tag
theUfSession.Modl.AskAllMembersOfSet(tempFeat.Tag, setFeatureTags, numFeatures)
Dim setFeatures As New List(Of Features.Feature)
if (tempFeat.name = "_2D_Drawing_Geometry" or tempFeat.name = "_2D_Zeichnungsgeometrie")
'get features from tags
For Each tempTag As Tag In setFeatureTags
	setFeatures.Add(Utilities.NXObjectManager.Get(tempTag))
Next
For Each setFeat As Features.Feature In setFeatures
REM MsgBox(setFeat.GetType.ToString)

If TypeOf (setFeat) Is Features.PointFeature Then
'get the point object from the feature
Dim nxObj() As NXObject = setFeat.GetEntities
For Each temppoint As Point In nxObj
Dim pobjects2(0) As NXOpen.DisplayableObject
pobjects2(0) = temppoint
displayModification1.NewLayer = PLayer
displayModification1.Apply(pobjects2)
pobjects1.Add(pobjects2(0))
next

else if TypeOf (setFeat) Is Features.TrimCurve2 Then
msgbox("success")
'get the point object from the feature
Dim nxObj() As NXObject = setFeat.GetEntities
For Each tempCurve As Curve In nxObj
Dim intobjects2(0) As NXOpen.DisplayableObject
intobjects2(0) = tempCurve
displayModification1.NewLayer = ILayer
displayModification1.Apply(intobjects2)
intobjects1.Add(intobjects2(0))
Next

REM else if TypeOf (setFeat) Is Features.IntersectionCurve Then
REM 'get the point object from the feature

REM if (intobjects1.count = 0)
REM Dim nxObj() As NXObject = setFeat.GetEntities
REM For Each tempCurve As Curve In nxObj
REM Dim intobjects2(0) As NXOpen.DisplayableObject
REM intobjects2(0) = tempCurve
REM displayModification1.NewLayer = ILayer
REM displayModification1.Apply(intobjects2)
REM intobjects1.Add(intobjects2(0))
REM Next
REM end if
		
end if
Next
else
'msgbox("NoObjects found.")				
end if
End If 
Next
'Points & Intersection Curves to 41 & 51 Layer



'Points & Intersection Curves to 41 & 51 Layer
For Each tempFeat As Features.Feature In workPart.Features
If TypeOf (tempFeat) Is Features.FeatureGroup Then
Dim numFeatures As Integer
Dim setFeatureTags() As Tag
theUfSession.Modl.AskAllMembersOfSet(tempFeat.Tag, setFeatureTags, numFeatures)
Dim setFeatures As New List(Of Features.Feature)
if (tempFeat.name = "_2D_Drawing_Geometry" or tempFeat.name = "_2D_Zeichnungsgeometrie")
'get features from tags
For Each tempTag As Tag In setFeatureTags
	setFeatures.Add(Utilities.NXObjectManager.Get(tempTag))
Next
For Each setFeat As Features.Feature In setFeatures
REM MsgBox(setFeat.GetType.ToString)

If TypeOf (setFeat) Is Features.IntersectionCurve Then
'get the point object from the feature

if (intobjects1.count = 0)
Dim nxObj() As NXObject = setFeat.GetEntities
For Each tempCurve As Curve In nxObj
Dim intobjects2(0) As NXOpen.DisplayableObject
intobjects2(0) = tempCurve
displayModification1.NewLayer = ILayer
displayModification1.Apply(intobjects2)
intobjects1.Add(intobjects2(0))
Next

end if
		
end if
Next
else
'msgbox("NoObjects found.")				
end if
End If 
Next
'Points & Intersection Curves to 41 & 51 Layer

'Add First Extrude to Final Part & Drawing Reference Set
Dim myReferenceSets As ReferenceSet()
myReferenceSets = workPart.GetAllReferenceSets()
Dim theReferenceSet As ReferenceSet = Nothing

'Try for adding reference set 
REM Try 
Const fprefset As String = "FINAL_PART"        
Const dwgrefset As String = "DRAWING"

For Each myRefSet As ReferenceSet In myReferenceSets		

If myRefSet.Name.ToUpper() = fprefset Then			
theReferenceSet = myRefSet

try
If refSetObjs IsNot Nothing Then
theReferenceSet.RemoveObjectsFromReferenceSet(theReferenceSet.AskAllDirectMembers())            
theReferenceSet.AddObjectsToReferenceSet(refSetObjs)
End If
Catch ex As NXException
REM msgbox("FinalPart: " & ex.message)
End Try

else if myRefSet.Name.ToUpper() = dwgrefset Then			
theReferenceSet = myRefSet	

Try
If refSetObjs IsNot Nothing AndAlso refSetObjs.length > 0 Then
theReferenceSet.RemoveObjectsFromReferenceSet(theReferenceSet.AskAllDirectMembers())         
theReferenceSet.AddObjectsToReferenceSet(refSetObjs)
end if
Catch ex As NXException
REM msgbox("Drawing: " & ex.message)
End Try

Try
'If bool > 0 Then       
theReferenceSet.AddObjectsToReferenceSet(outobjects1)
'end if
Catch ex As NXException
REM msgbox("Drawing-Flamecut: " & ex.message)
End Try

Try
If pobjects1 IsNot Nothing AndAlso pobjects1.Count > 0 Then
dim ppobjects1() as NXOpen.DisplayableObject = pobjects1.ToArray()			
theReferenceSet.AddObjectsToReferenceSet(ppobjects1)
end if
Catch ex As NXException
REM msgbox("Drawing-Points: " & ex.message)
End Try

Try
If intobjects1 IsNot Nothing AndAlso intobjects1.Count > 0 Then
dim iintobjects1() as NXOpen.DisplayableObject = intobjects1.ToArray()			
theReferenceSet.AddObjectsToReferenceSet(iintobjects1)
End If
Catch ex As NXException
REM msgbox("Drawing-Curves: " & ex.message)
End Try

end if

Next
REM Catch ex As NXException
REM msgbox("Add Refset: " & ex.message)
REM End Try
'Try for adding reference set

'Add First Extrude to Final Part & Drawing Reference Set

displayModification1.Dispose()
theSession.CleanUpFacetedFacesAndEdges()

workPart.Preferences.Modeling.CutViewUpdateDelayed = True
theSession.UpdateManager.InterpartDelay = False	

pobjects1.clear()
intobjects1.clear()

Dim partSaveStatus1 As NXOpen.PartSaveStatus = Nothing
partSaveStatus1 = workPart.Save(NXOpen.BasePart.SaveComponents.False, NXOpen.BasePart.CloseAfterSave.True)
partLoadStatus1.Dispose()

else
'msgbox("Mirrored Part.")
end if

Next

Dim part2 As NXOpen.Part = CType(theSession.Parts.FindObject(mainassy), NXOpen.Part)

Dim partLoadStatus2 As NXOpen.PartLoadStatus
Dim status2 As NXOpen.PartCollection.SdpsStatus
status2 = theSession.Parts.SetDisplay(part2, False, True, partLoadStatus2)

workPart = theSession.Parts.Work
displayPart = theSession.Parts.Display

Dim componentsToOpen1(0) As NXOpen.Assemblies.Component
componentsToOpen1(0) = workPart.ComponentAssembly.RootComponent
Dim openStatus1() As NXOpen.Assemblies.ComponentAssembly.OpenComponentStatus
Dim partLoadStatus3 As NXOpen.PartLoadStatus
partLoadStatus3 = workPart.ComponentAssembly.OpenComponents(NXOpen.Assemblies.ComponentAssembly.OpenOption.WholeAssembly, componentsToOpen1, openStatus1)

partLoadStatus3.Dispose()

theSession.UpdateManager.InterpartDelay = True

'Close Main Try		
Catch e As Exception
REM theSession.ListingWindow.WriteLine("Failed: " & e.ToString)
UI.GetUI.NXMessageBox.Show("Message", NXMessageBox.DialogType.Information,"Please Open Assembly File and Try Again.")
REM msgbox("Main try " & e.message)
End Try
'lw.Close

End Sub
 
'**********************************************************
Sub reportComponentChildren( ByVal comp As Component, _
ByVal indent As Integer)

For Each child As Component In comp.GetChildren()
Dim chname as string = child.DisplayName()

if (chname.contains("F5") And chname.contains("GEO") And Not chname.contains("Adapter"))
REM msgbox(chname.substring(17,1))
if((chname.substring(17,1)<>"3") And (chname.substring(17,1)<>"4") And (chname.substring(17,1)<>"5"))
Complist.add(child)	
Complistname.add(child.DisplayName())
end if
end if
reportComponentChildren(child, indent + 1)
Next
End Sub
'**********************************************************
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        Return Session.LibraryUnloadOption.Immediately
    End Function
'**********************************************************
 
End Module