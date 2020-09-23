VERSION 5.00
Begin VB.Form frmX 
   AutoRedraw      =   -1  'True
   Caption         =   "Welcome To Visual X"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   Icon            =   "frmX.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5235
      Left            =   60
      ScaleHeight     =   5205
      ScaleWidth      =   6345
      TabIndex        =   1
      Top             =   60
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Obj Edit"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   4920
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   180
      Top             =   180
   End
End
Attribute VB_Name = "frmX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim poly As vx_Polygon
Dim sides(6) As vx_PolySide

Dim poly2 As vx_Polygon
Dim p2sides(5) As vx_PolySide

Dim vx_Dimond As vx_Polygon
Dim vx_DimondSides(5) As vx_PolySide

Private vx_Ending As Boolean
Private vx_FPS As Long
Private vx_Edit As Boolean

Private Sub Command1_Click()
vx_Edit = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

vx_Ending = True

End Sub

Private Sub Form_Load()

InitVisualX

modVisualX.Canvas = Picture1

'Auto Generated Using The Visual X Editor

poly.sides = 7
poly.origin.X = 200
poly.origin.Y = 200
poly.origin.Z = 0
sides(0).c_points = 7
sides(0).points(0).X = -30
sides(0).points(0).Y = -30
sides(0).points(0).Z = 0
sides(0).points(1).X = -5
sides(0).points(1).Y = 20
sides(0).points(1).Z = 0
sides(0).points(2).X = 5
sides(0).points(2).Y = 20
sides(0).points(2).Z = 0
sides(0).points(3).X = 30
sides(0).points(3).Y = -30
sides(0).points(3).Z = 0
sides(0).points(4).X = 20
sides(0).points(4).Y = -30
sides(0).points(4).Z = 0
sides(0).points(5).X = 0
sides(0).points(5).Y = 10
sides(0).points(5).Z = 0
sides(0).points(6).X = -20
sides(0).points(6).Y = -30
sides(0).points(6).Z = 0
sides(1).c_points = 4
sides(1).points(0).X = 30
sides(1).points(0).Y = 0
sides(1).points(0).Z = 0
sides(1).points(1).X = 30
sides(1).points(1).Y = 20
sides(1).points(1).Z = 0
sides(1).points(2).X = 40
sides(1).points(2).Y = 20
sides(1).points(2).Z = 0
sides(1).points(3).X = 40
sides(1).points(3).Y = 0
sides(1).points(3).Z = 0
sides(2).c_points = 11
sides(2).points(0).X = 45
sides(2).points(0).Y = 0
sides(2).points(0).Z = 0
sides(2).points(1).X = 45
sides(2).points(1).Y = 13
sides(2).points(1).Z = 0
sides(2).points(2).X = 55
sides(2).points(2).Y = 13
sides(2).points(2).Z = 0
sides(2).points(3).X = 45
sides(2).points(3).Y = 13
sides(2).points(3).Z = 0
sides(2).points(4).X = 45
sides(2).points(4).Y = 20
sides(2).points(4).Z = 0
sides(2).points(5).X = 60
sides(2).points(5).Y = 20
sides(2).points(5).Z = 0
sides(2).points(6).X = 60
sides(2).points(6).Y = 7
sides(2).points(6).Z = 0
sides(2).points(7).X = 50
sides(2).points(7).Y = 7
sides(2).points(7).Z = 0
sides(2).points(8).X = 60
sides(2).points(8).Y = 7
sides(2).points(8).Z = 0
sides(2).points(9).X = 60
sides(2).points(9).Y = 0
sides(2).points(9).Z = 0
sides(2).points(10).X = 45
sides(2).points(10).Y = 0
sides(2).points(10).Z = 0
sides(3).c_points = 8
sides(3).points(0).X = 65
sides(3).points(0).Y = 0
sides(3).points(0).Z = 0
sides(3).points(1).X = 65
sides(3).points(1).Y = 20
sides(3).points(1).Z = 0
sides(3).points(2).X = 80
sides(3).points(2).Y = 20
sides(3).points(2).Z = 0
sides(3).points(3).X = 80
sides(3).points(3).Y = 0
sides(3).points(3).Z = 0
sides(3).points(4).X = 72
sides(3).points(4).Y = 0
sides(3).points(4).Z = 0
sides(3).points(5).X = 72
sides(3).points(5).Y = 15
sides(3).points(5).Z = 0
sides(3).points(6).X = 72
sides(3).points(6).Y = 0
sides(3).points(6).Z = 0
sides(3).points(7).X = 65
sides(3).points(7).Y = 0
sides(3).points(7).Z = 0
sides(4).c_points = 8
sides(4).points(0).X = 90
sides(4).points(0).Y = 0
sides(4).points(0).Z = 0
sides(4).points(1).X = 85
sides(4).points(1).Y = 20
sides(4).points(1).Z = 0
sides(4).points(2).X = 95
sides(4).points(2).Y = 20
sides(4).points(2).Z = 0
sides(4).points(3).X = 95
sides(4).points(3).Y = 10
sides(4).points(3).Z = 0
sides(4).points(4).X = 95
sides(4).points(4).Y = 20
sides(4).points(4).Z = 0
sides(4).points(5).X = 105
sides(4).points(5).Y = 20
sides(4).points(5).Z = 0
sides(4).points(6).X = 100
sides(4).points(6).Y = 0
sides(4).points(6).Z = 0
sides(4).points(7).X = 90
sides(4).points(7).Y = 0
sides(4).points(7).Z = 0
sides(5).c_points = 7
sides(5).points(0).X = 110
sides(5).points(0).Y = 0
sides(5).points(0).Z = 0
sides(5).points(1).X = 110
sides(5).points(1).Y = 20
sides(5).points(1).Z = 0
sides(5).points(2).X = 125
sides(5).points(2).Y = 20
sides(5).points(2).Z = 0
sides(5).points(3).X = 125
sides(5).points(3).Y = 15
sides(5).points(3).Z = 0
sides(5).points(4).X = 115
sides(5).points(4).Y = 15
sides(5).points(4).Z = 0
sides(5).points(5).X = 115
sides(5).points(5).Y = 0
sides(5).points(5).Z = 0
sides(5).points(6).X = 110
sides(5).points(6).Y = 0
sides(5).points(6).Z = 0
sides(6).c_points = 13
sides(6).points(0).X = 130
sides(6).points(0).Y = 0
sides(6).points(0).Z = 0
sides(6).points(1).X = 135
sides(6).points(1).Y = 10
sides(6).points(1).Z = 0
sides(6).points(2).X = 130
sides(6).points(2).Y = 20
sides(6).points(2).Z = 0
sides(6).points(3).X = 135
sides(6).points(3).Y = 20
sides(6).points(3).Z = 0
sides(6).points(4).X = 140
sides(6).points(4).Y = 12
sides(6).points(4).Z = 0
sides(6).points(5).X = 145
sides(6).points(5).Y = 20
sides(6).points(5).Z = 0
sides(6).points(6).X = 150
sides(6).points(6).Y = 20
sides(6).points(6).Z = 0
sides(6).points(7).X = 145
sides(6).points(7).Y = 10
sides(6).points(7).Z = 0
sides(6).points(8).X = 150
sides(6).points(8).Y = 0
sides(6).points(8).Z = 0
sides(6).points(9).X = 145
sides(6).points(9).Y = 0
sides(6).points(9).Z = 0
sides(6).points(10).X = 140
sides(6).points(10).Y = 8
sides(6).points(10).Z = 0
sides(6).points(11).X = 135
sides(6).points(11).Y = 0
sides(6).points(11).Z = 0
sides(6).points(12).X = 130
sides(6).points(12).Y = 0
sides(6).points(12).Z = 0

sides(0).fBrush = CreateSolidBrush(RGB(127, 127, 127))
sides(1).fBrush = sides(0).fBrush
sides(2).fBrush = sides(0).fBrush
sides(3).fBrush = sides(0).fBrush
sides(4).fBrush = sides(0).fBrush
sides(5).fBrush = sides(0).fBrush
sides(6).fBrush = CreateSolidBrush(vbRed)

'Auto Generated Using The Visual X Editor
vx_Dimond.sides = 6
vx_Dimond.origin.X = 40
vx_Dimond.origin.Y = 40
vx_Dimond.origin.Z = 0
vx_DimondSides(0).c_points = 4
vx_DimondSides(0).points(0).X = -20
vx_DimondSides(0).points(0).Y = -20
vx_DimondSides(0).points(0).Z = 20
vx_DimondSides(0).points(1).X = 20
vx_DimondSides(0).points(1).Y = -20
vx_DimondSides(0).points(1).Z = 20
vx_DimondSides(0).points(2).X = 20
vx_DimondSides(0).points(2).Y = 20
vx_DimondSides(0).points(2).Z = 20
vx_DimondSides(0).points(3).X = -20
vx_DimondSides(0).points(3).Y = 20
vx_DimondSides(0).points(3).Z = 20
vx_DimondSides(1).c_points = 4
vx_DimondSides(1).points(0).X = -20
vx_DimondSides(1).points(0).Y = -20
vx_DimondSides(1).points(0).Z = -20
vx_DimondSides(1).points(1).X = 20
vx_DimondSides(1).points(1).Y = -20
vx_DimondSides(1).points(1).Z = -20
vx_DimondSides(1).points(2).X = 20
vx_DimondSides(1).points(2).Y = 20
vx_DimondSides(1).points(2).Z = -20
vx_DimondSides(1).points(3).X = -20
vx_DimondSides(1).points(3).Y = 20
vx_DimondSides(1).points(3).Z = -20
vx_DimondSides(2).c_points = 4
vx_DimondSides(2).points(0).X = -20
vx_DimondSides(2).points(0).Y = -20
vx_DimondSides(2).points(0).Z = -20
vx_DimondSides(2).points(1).X = -20
vx_DimondSides(2).points(1).Y = -20
vx_DimondSides(2).points(1).Z = 20
vx_DimondSides(2).points(2).X = 20
vx_DimondSides(2).points(2).Y = -20
vx_DimondSides(2).points(2).Z = 20
vx_DimondSides(2).points(3).X = 20
vx_DimondSides(2).points(3).Y = -20
vx_DimondSides(2).points(3).Z = -20
vx_DimondSides(3).c_points = 4
vx_DimondSides(3).points(0).X = -20
vx_DimondSides(3).points(0).Y = -20
vx_DimondSides(3).points(0).Z = 20
vx_DimondSides(3).points(1).X = -20
vx_DimondSides(3).points(1).Y = -20
vx_DimondSides(3).points(1).Z = -20
vx_DimondSides(3).points(2).X = -20
vx_DimondSides(3).points(2).Y = 20
vx_DimondSides(3).points(2).Z = -20
vx_DimondSides(3).points(3).X = -20
vx_DimondSides(3).points(3).Y = 20
vx_DimondSides(3).points(3).Z = 20
vx_DimondSides(4).c_points = 4
vx_DimondSides(4).points(0).X = 20
vx_DimondSides(4).points(0).Y = -20
vx_DimondSides(4).points(0).Z = 20
vx_DimondSides(4).points(1).X = 20
vx_DimondSides(4).points(1).Y = -20
vx_DimondSides(4).points(1).Z = -20
vx_DimondSides(4).points(2).X = 20
vx_DimondSides(4).points(2).Y = 20
vx_DimondSides(4).points(2).Z = -20
vx_DimondSides(4).points(3).X = 20
vx_DimondSides(4).points(3).Y = 20
vx_DimondSides(4).points(3).Z = 20
vx_DimondSides(5).c_points = 4
vx_DimondSides(5).points(0).X = -20
vx_DimondSides(5).points(0).Y = 20
vx_DimondSides(5).points(0).Z = 20
vx_DimondSides(5).points(1).X = -20
vx_DimondSides(5).points(1).Y = 20
vx_DimondSides(5).points(1).Z = -20
vx_DimondSides(5).points(2).X = 20
vx_DimondSides(5).points(2).Y = 20
vx_DimondSides(5).points(2).Z = -20
vx_DimondSides(5).points(3).X = 20
vx_DimondSides(5).points(3).Y = 20
vx_DimondSides(5).points(3).Z = 20


vx_DimondSides(0).fBrush = CreateSolidBrush(vbRed)
vx_DimondSides(1).fBrush = vx_DimondSides(0).fBrush
vx_DimondSides(2).fBrush = vx_DimondSides(0).fBrush
vx_DimondSides(3).fBrush = vx_DimondSides(0).fBrush
vx_DimondSides(4).fBrush = vx_DimondSides(0).fBrush
vx_DimondSides(5).fBrush = vx_DimondSides(0).fBrush


'DrawPolygon poly, sides

DrawLoop

End Sub


Public Sub DrawLoop()

Me.Show
DoEvents

again:

poly.angle.X = poly.angle.X + 3
poly.angle.Y = poly.angle.Y + 1
poly.angle.Z = poly.angle.Z + 2

If poly.angle.X = 360 Then poly.angle.X = 0
If poly.angle.Y = 360 Then poly.angle.Y = 0
If poly.angle.Z = 360 Then poly.angle.Z = 0
    
vx_Dimond.angle.Y = vx_Dimond.angle.Y + 1

If vx_Dimond.angle.Y = 360 Then vx_Dimond.angle.Y = 0

vx_Dimond.angle.X = vx_Dimond.angle.X + 1

If vx_Dimond.angle.X = 360 Then vx_Dimond.angle.X = 0

vx_Dimond.angle.Z = vx_Dimond.angle.Z + 1

If vx_Dimond.angle.Z = 360 Then vx_Dimond.angle.Z = 0
    
Picture1.Cls

DrawPolygon poly, sides
DrawPolygon vx_Dimond, vx_DimondSides

Picture1.Refresh
DoEvents

vx_FPS = vx_FPS + 1

If vx_Edit = True Then
    frmEdit.Show
    Exit Sub
End If

If Not vx_Ending Then GoTo again

End

End Sub

Private Sub Timer1_Timer()
Me.Caption = "Welcome To Visual X - " & vx_FPS
vx_FPS = 0
End Sub
