'Program to Create Part with Axis systems from Excel Input
'By GopinadhGVS on Dt.27-11-2020

Imports NXOpen, NXOpenUI
Imports System, System.Drawing.Color, System.Windows.Forms
Imports System.Collections.Generic
Imports System.Collections
Imports NXOpen.UF
Imports System.IO
Imports System.Globalization


Module NXJournal

Public loc As string
Public Fldr As string
public uv as new list (of string)
Public fpath As String
Public pname As String
Public fpaths() As String
Public part1 As NXOpen.Part

 

Sub Main()

Dim theSession As Session = Session.GetSession 'A variable to hold the NX Session
Dim markId1 As NXOpen.Session.UndoMarkId = Nothing
markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.visible, "Create CSYS")


Dim displayPart As NXOpen.Part = theSession.Parts.Display
' theSession = Session.GetSession 'Get the NX Session
    Dim theUfSession As UFSession = UFSession.GetUFSession  
    Dim lw As ListingWindow = theSession.ListingWindow
    Dim workPart As NXOpen.Part = theSession.Parts.Work
	
'Form Launch

Dim f1 As New Form1           
f1.ShowDialog()

Dim path as string
path = Loc

If Loc = "" Then Exit Sub

Dim EXCEL = CreateObject("Excel.Application")
EXCEL.Visible = False
Dim Doc = EXCEL.Workbooks.Open(path, ReadOnly:=True)
Dim Sheets = EXCEL.Sheets
Dim Sheet = Doc.Sheets.Item(1)
Dim pointCounter As Integer
For pointCounter = 1 To Sheet.UsedRange.Rows.Count - 1
pname = Sheet.Cells(2,1).Value
Next
Doc.Close()
EXCEL.Quit()

If pname = ""
MsgBox("Invalid Partname from Excel., Check & Try again.!!!", vbOKOnly + vbCritical, "Try Again")
Exit Sub
else
end if


Dim markId10 As NXOpen.Session.UndoMarkId = Nothing
markId10 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Start")

Dim fileNew1 As NXOpen.FileNew = Nothing
fileNew1 = theSession.Parts.FileNew()

theSession.SetUndoMarkName(markId10, "New Dialog")

Dim markId2 As NXOpen.Session.UndoMarkId = Nothing
markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "New")

theSession.DeleteUndoMark(markId2, Nothing)

Dim markId3 As NXOpen.Session.UndoMarkId = Nothing
markId3 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "New")

fileNew1.TemplateFileName = "daimler_startpart_model.prt"

fileNew1.UseBlankTemplate = False

fileNew1.ApplicationName = "ModelTemplate"

fileNew1.Units = NXOpen.Part.Units.Millimeters

fileNew1.RelationType = ""

fileNew1.UsesMasterModel = "No"

fileNew1.TemplateType = NXOpen.FileNewTemplateType.Item

fileNew1.TemplatePresentationName = "Model (NX Geometry part)"

fileNew1.ItemType = ""

fileNew1.Specialization = ""

fileNew1.SetCanCreateAltrep(False)

fileNew1.NewFileName = "C:\temp\" + pname + ".prt"

fileNew1.MasterFileName = ""

fileNew1.MakeDisplayedPart = True

fileNew1.DisplayPartOption = NXOpen.DisplayPartOption.AllowAdditional

Try

Dim nXObject10 As NXOpen.NXObject = Nothing
nXObject10 = fileNew1.Commit()

Catch ex As NXException

MsgBox("Part With Same filename Exits.", vbOKOnly + vbCritical, "Try Again")
Exit Sub
End Try

workPart = theSession.Parts.Work ' _model2
displayPart = theSession.Parts.Display ' _model2
theSession.DeleteUndoMark(markId3, Nothing)

fileNew1.Destroy()


Dim X2, Y2, Z2,AX2,AY2,AZ2 As Double
dim csysname as string

Dim Doc1 = EXCEL.Workbooks.Open(path, ReadOnly:=True)
Dim Sheets1 = EXCEL.Sheets
Dim Sheet1 = Doc1.Sheets.Item(1)
For pointCounter = 2 To Sheet1.UsedRange.Rows.Count - 1

If (IsNumeric(Sheet1.Cells((pointCounter + 1), 2).Value))
else
MsgBox("Invalid Coordinate.,Check & Try again.!!!", vbOKOnly + vbCritical, "Try Again")
Exit Sub
end if

If (IsNumeric(Sheet1.Cells((pointCounter + 1), 3).Value))
else
MsgBox("Invalid Coordinate.,Check & Try again.!!!", vbOKOnly + vbCritical, "Try Again")
Exit Sub
end if

If (IsNumeric(Sheet1.Cells((pointCounter + 1), 4).Value))
else
MsgBox("Invalid Coordinate.,Check & Try again.!!!", vbOKOnly + vbCritical, "Try Again")
Exit Sub
end if

If (IsNumeric(Sheet1.Cells((pointCounter + 1), 5).Value))
else
MsgBox("Invalid Coordinate.,Check & Try again.!!!", vbOKOnly + vbCritical, "Try Again")
Exit Sub
end if

If (IsNumeric(Sheet1.Cells((pointCounter + 1), 6).Value))
else
MsgBox("Invalid Coordinate.,Check & Try again.!!!", vbOKOnly + vbCritical, "Try Again")
Exit Sub
end if

If (IsNumeric(Sheet1.Cells((pointCounter + 1), 7).Value))
else
MsgBox("Invalid Coordinate.,Check & Try again.!!!", vbOKOnly + vbCritical, "Try Again")
Exit Sub
end if

csysname = Sheet1.Cells((pointCounter + 1), 1).Value
X2 = Sheet1.Cells((pointCounter + 1), 2).Value
Y2 = Sheet1.Cells((pointCounter + 1), 3).Value
Z2 = Sheet1.Cells((pointCounter + 1), 4).Value
AX2 = Sheet1.Cells((pointCounter + 1), 5).Value
AY2 = Sheet1.Cells((pointCounter + 1), 6).Value
AZ2 = Sheet1.Cells((pointCounter + 1), 7).Value


Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing
Dim datumCsysBuilder1 As NXOpen.Features.DatumCsysBuilder = Nothing
datumCsysBuilder1 = workPart.Features.CreateDatumCsysBuilder(nullNXOpen_Features_Feature)
Dim pnt As Point3D
Dim NewPoint As Point
Dim Xp, Yp, Zp As Double
Xp = x2
Yp = y2
Zp = z2
pnt = New Point3D(Xp, Yp, Zp)

        Dim oldOrigin As Point3d = displayPart.WCS.Origin
        Dim oldOrientation As CartesianCoordinateSystem = displayPart.WCS.CoordinateSystem 
        'move WCS origin 10 units in +X (absolute) direction
        'Dim newOrigin As New Point3d(oldOrigin.X + 10, oldOrigin.Y, oldOrigin.Z)
 
        'move WCS, the .SetOriginAndMatrix method does not save the old WCS
        displayPart.WCS.SetOriginAndMatrix(pnt, oldOrientation.Orientation.Element)

        displayPart.WCS.Rotate(WCS.Axis.ZAxis, AZ2)
        displayPart.WCS.Rotate(WCS.Axis.YAxis, AY2)
        displayPart.WCS.Rotate(WCS.Axis.XAxis, AX2)


'----------------------------------------------------------------------------------------
        Dim absXform As Xform = displayPart.Xforms.CreateXform(SmartObject.UpdateOption.WithinModeling, 1)        
        Dim absCsys As CartesianCoordinateSystem = displayPart.CoordinateSystems.CreateCoordinateSystem(absXform, SmartObject.UpdateOption.WithinModeling)
        Dim csys2 As CartesianCoordinateSystem = displayPart.WCS.SetCoordinateSystem(absCsys)

Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = csys2
Dim originOffset1 As NXOpen.Vector3d = New NXOpen.Vector3d(0.0, 0.0, 0.0)
Dim trasformMatrix1 As NXOpen.Matrix3x3 = Nothing
trasformMatrix1.Xx = 1.0
trasformMatrix1.Xy = 0.0
trasformMatrix1.Xz = 0.0
trasformMatrix1.Yx = 0.0
trasformMatrix1.Yy = 1.0
trasformMatrix1.Yz = 0.0
trasformMatrix1.Zx = 0.0
trasformMatrix1.Zy = 0.0
trasformMatrix1.Zz = 1.0
Dim xform1 As NXOpen.Xform = Nothing
xform1 = workPart.Xforms.CreateXformByDynamicOffset(cartesianCoordinateSystem1, originOffset1, trasformMatrix1, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)

Dim cartesianCoordinateSystem2 As NXOpen.CartesianCoordinateSystem = Nothing
cartesianCoordinateSystem2 = workPart.CoordinateSystems.CreateCoordinateSystem(xform1, NXOpen.SmartObject.UpdateOption.WithinModeling)

datumCsysBuilder1.Csys = cartesianCoordinateSystem2
datumCsysBuilder1.DisplayScaleFactor = 1.0
Dim nXObject1 As NXOpen.NXObject = Nothing
nXObject1 = datumCsysBuilder1.Commit()

Dim datumCsys1 As NXOpen.Features.DatumCsys = CType(nXObject1, NXOpen.Features.DatumCsys)

datumCsys1.SetName(csysname)

datumCsysBuilder1.Destroy()

Dim objects10(0) As NXOpen.TaggedObject

objects10(0) = csys2

Dim nErrs1 As Integer = Nothing
nErrs1 = theSession.UpdateManager.AddObjectsToDeleteList(objects10)

Dim notifyOnDelete2 As Boolean = Nothing
notifyOnDelete2 = theSession.Preferences.Modeling.NotifyOnDelete

Dim nErrs2 As Integer = Nothing
nErrs2 = theSession.UpdateManager.DoUpdate(markId1)

theSession.DeleteUndoMark(markId1, Nothing)

Next
workPart.ModelingViews.WorkView.Fit()
Doc1.Close()
EXCEL.Quit()


End Sub


End Module



'Form Working
Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        
        NXJournal.loc = Box.Text
        'NXJournal.Splayer = Box1.Text

        if (box.text <> "")         
        Me.Close()      
        else  
		MessageBox.Show("Please Select Valid Path for Excel file & Try Again.", "Excel Path Error", _
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
		end if
		
		REM if (box1.text <> "")         
        REM Me.Close()      
        REM else        
		REM MessageBox.Show("Invalid Layer Number, Try Again.", "Layer Number Error", _
                        REM MessageBoxButtons.OK, MessageBoxIcon.Information)
        REM end if
      
    End Sub 

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    
		MessageBox.Show("Exiting the Program.", "Exit", _
                        MessageBoxButtons.OK, MessageBoxIcon.Information)		
        Me.Close()
        exit sub
      
    End Sub

    
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim OpenFileDialog1 As New OpenFileDialog
		openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
       	openFileDialog1.FilterIndex = 0
       	openFileDialog1.Title = "Select Excel file containing Points information"
       	openFileDialog1.RestoreDirectory = True
       	openFileDialog1.Autoupgradeenabled = True
        If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
            Box.Text = OpenFileDialog1.FileName
        End If
      
   End Sub
           
      
End Class

'Form Design
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

   'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
            Me.lbl = New System.Windows.Forms.Label()
            Me.lb2 = New System.Windows.Forms.Label()
            Me.lb1 = New System.Windows.Forms.Label()
            Me.box = New System.Windows.Forms.TextBox()
            Me.box2 = New System.Windows.Forms.TextBox()
            Me.box1 = New System.Windows.Forms.TextBox()
            'e.com = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(200, 50)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(66, 27)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Run"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(300, 50)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(66, 27)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        
        
        
        Me.Button3.Location = New System.Drawing.Point(500, 10)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(66, 27)
        Me.Button3.TabIndex = 0
        Me.Button3.Text = "Browse"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Form1
        
        lbl.Location = New System.Drawing.Point(10, 10) 'set your location
            lbl.AutoSize = True
            lbl.Text = "Browse for Excel File" 'set the text for your label
            Me.Controls.Add(lbl)
            
            
            lb2.Location = New System.Drawing.Point(10, 50) 'set your location
            lb2.AutoSize = True
            lb2.Text = "Sphere Layer Number" 'set the text for your label
            'Me.Controls.Add(lb2)
           
            box.Location = New System.Drawing.Point(170,10)
            Box.Width = 300
            box.Text = ""
            Me.Controls.Add(box)
			
			box1.Location = New System.Drawing.Point(170,50)
            Box1.Width = 50
            box1.Text = "1"
            'Me.Controls.Add(box1)
        
        'Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        'Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            'Me.BackColor = System.Drawing.Color.SteelBlue
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        'Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.ClientSize = New System.Drawing.Size(600, 100)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Axis_System_Generator."
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
      Friend WithEvents lbl As System.Windows.Forms.Label
      Friend WithEvents lb2 As System.Windows.Forms.Label
      Friend WithEvents lb1 As System.Windows.Forms.Label
      Friend WithEvents box As System.Windows.Forms.TextBox
      Friend WithEvents box2 As System.Windows.Forms.TextBox
      Friend WithEvents box1 As System.Windows.Forms.TextBox
      'Friend WithEvents Com As System.Windows.Forms.ComboBox
End Class