Option Strict Off
Imports System
Imports System.Windows.Forms
Imports NXOpen
Imports System.Windows.Controls


Module Module1
 
    Sub Main()
 
        Dim theSession As Session = Session.GetSession()
        'If IsNothing(theSession.Parts.BaseWork) Then
            'active part required
            'Return
        'End If
 
        'Dim workPart As Part = theSession.Parts.Work
        Dim lw As ListingWindow = theSession.ListingWindow
        lw.Open()
        
 
 Dim openFileDialog1 As New OpenFileDialog()

 
       'openFileDialog1.InitialDirectory = "Q:\daten\DAIMLER\MFA2_Door\"
       'openFileDialog1.Filter = "prt files (*.prt)|*.prt|STP files (*.stp)|*.stp|STEP files (*.step)|*.step|All files (*.*)|*.*"
        openFileDialog1.Filter = "prt files (*.prt)|*.prt|DWG files (**DWG*)|**DWG*|STP files (*.stp)|*.stp|STEP files (*.step)|*.step|All files (*.*)|*.*"
       	openFileDialog1.FilterIndex = 0
       	openFileDialog1.Title = "Multi Open Dialog box"
       	openFileDialog1.RestoreDirectory = True
       	openFileDialog1.Autoupgradeenabled = True
        openFileDialog1.Multiselect = True

 
        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                For Each prt As String In openFileDialog1.FileNames
                    'lw.WriteLine(prt)
                    Dim basePart1 As BasePart
                    Dim partLoadStatus1 As PartLoadStatus
                    basePart1 = theSession.Parts.OpenBaseDisplay(prt, partLoadStatus1)
 
                    partLoadStatus1.Dispose()
 
                Next
            Catch Ex As Exception
                lw.WriteLine("Error: " & Ex.Message)
            End Try
        End If
 
        lw.Close()
 
    End Sub
 
 	
 
 Public Function GetUnloadOption(ByVal dummy As String) As Integer
 
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
 
    End Function
 
End Module





