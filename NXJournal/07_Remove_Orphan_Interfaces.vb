Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.Features
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports NXOpen.Assemblies
 
 
Module Module1
 
    Sub Main()
 
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim lw As ListingWindow = theSession.ListingWindow
        lw.Open()

'-------------------------------------Undo Mark Start--------------------------------------

Dim markId2 As NXOpen.Session.UndoMarkId = Nothing
markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Orphan Product Interface Remover")

'-------------------------------------Undo Mark End--------------------------------------

'----------------Check For Active part  if Not available Close the Program----------------
        If IsNothing(theSession.Parts.BaseWork) Then

        MessageBox.Show("Active Part not found")

        Return
        End If


lw.WriteLine("**********************************************************************")
lw.WriteLine("                 Orphan Product Interface Remover                     ")
lw.WriteLine("**********************************************************************")
lw.WriteLine(" ")



Dim objectBuilder1 As NXOpen.Assemblies.ProductInterface.ObjectBuilder = Nothing
objectBuilder1 = workPart.ProductInterface.CreateObjectBuilderWithVersion(2)

Dim prodIntObjs1() As NXOpen.Assemblies.ProductInterface.InterfaceObject
prodIntObjs1 = objectBuilder1.QueryProductInterfaceObjects(workPart)

dim j as integer = 0
Dim Statlistbefore As New List (of String)
Dim Statlistafter As New List (of String)

    for each myobj As NXOpen.Assemblies.ProductInterface.InterfaceObject In prodIntObjs1

        myobj = prodIntObjs1(j)
        ' Statlistbefore.Add("Product Interface : " & myobj.Name & " Status is : " & myobj.GetProductInterfaceObjectStatus)
        ' Statlistbefore.Add("Product Interface : " & myobj.GetProductInterfaceObjectType & " Status is : " & myobj.GetProductInterfaceObjectStatus)
        Statlistbefore.Add("Product Interface : " & myobj.GetProductInterfaceObjectType & " Status is : " & myobj.GetProductInterfaceObjectStatus)
        
        

            If myobj.GetProductInterfaceObjectStatus = "Orphan"


                myobj.RemoveProductInterfaceObject()

                Dim objects1(-1) As NXOpen.TaggedObject
                Dim nErrs1 As Integer = Nothing
                nErrs1 = theSession.UpdateManager.AddObjectsToDeleteList(objects1)

                Dim notifyOnDelete2 As Boolean = Nothing
                notifyOnDelete2 = theSession.Preferences.Modeling.NotifyOnDelete

                Dim nErrs2 As Integer = Nothing
                nErrs2 = theSession.UpdateManager.DoUpdate(markId2)

            End If

        j=j+1


        ' Statlistbefore.Add("Product Interface : " & myobj.Name & " Status is : " & myobj.GetProductInterfaceObjectStatus)
        Statlistafter.Add("Product Interface : " & myobj.GetProductInterfaceObjectType & " Status is : " & myobj.GetProductInterfaceObjectStatus)

    Next


                lw.writeline("Interface Status Before")
                lw.writeline("_______________________")
        For Each beforename As String In Statlistbefore
                lw.writeline(beforename)
        Next
lw.writeline("")
lw.writeline("")
                lw.writeline("Interface Status After")
                lw.writeline("_______________________")
        For Each Aftername As String In Statlistafter
                lw.writeline(Aftername)
        Next




 
 
    End Sub


    Public Function GetUnloadOption(ByVal dummy As String) As Integer
 
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
 
    End Function
 
End Module


