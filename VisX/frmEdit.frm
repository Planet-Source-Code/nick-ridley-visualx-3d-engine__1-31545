VERSION 5.00
Begin VB.Form frmEdit 
   Caption         =   "Visual X Object Editor"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5040
      Top             =   6180
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Generate"
      Height          =   495
      Left            =   8760
      TabIndex        =   26
      Top             =   6060
      Width           =   3255
   End
   Begin VB.TextBox Code 
      Height          =   6015
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Text            =   "frmEdit.frx":0000
      Top             =   0
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4500
      Top             =   6180
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Auto Spin"
      Height          =   315
      Left            =   3480
      TabIndex        =   24
      Top             =   5820
      Width           =   1455
   End
   Begin VB.HScrollBar SZ 
      Height          =   195
      Left            =   120
      Max             =   359
      TabIndex        =   23
      Top             =   6360
      Width           =   2595
   End
   Begin VB.HScrollBar SY 
      Height          =   195
      Left            =   120
      Max             =   359
      TabIndex        =   22
      Top             =   6120
      Width           =   2595
   End
   Begin VB.HScrollBar SX 
      Height          =   195
      Left            =   120
      Max             =   359
      TabIndex        =   21
      Top             =   5880
      Width           =   2595
   End
   Begin VB.TextBox txtOZ 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7740
      TabIndex        =   20
      Text            =   "0"
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox txtOY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6780
      TabIndex        =   19
      Text            =   "0"
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox txtOX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5820
      TabIndex        =   18
      Text            =   "0"
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox txtP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7740
      TabIndex        =   15
      Text            =   "1"
      Top             =   1680
      Width           =   795
   End
   Begin VB.TextBox txtZ 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   7740
      TabIndex        =   13
      Text            =   "0"
      Top             =   2280
      Width           =   795
   End
   Begin VB.TextBox txtY 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   6720
      TabIndex        =   12
      Text            =   "0"
      Top             =   2280
      Width           =   795
   End
   Begin VB.TextBox txtX 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   5700
      TabIndex        =   11
      Text            =   "0"
      Top             =   2280
      Width           =   795
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Draw"
      Height          =   495
      Left            =   7260
      TabIndex        =   10
      Top             =   6060
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "< "
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtNoSides 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      TabIndex        =   5
      Text            =   "1"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtVBName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      TabIndex        =   3
      Text            =   "vx_Object"
      Top             =   360
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5715
      Left            =   60
      ScaleHeight     =   5655
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   60
      Width           =   4875
   End
   Begin VB.Label lblFPS 
      Caption         =   "--"
      Height          =   195
      Left            =   3660
      TabIndex        =   27
      Top             =   6180
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "FPS Rate:"
      Height          =   255
      Left            =   2820
      TabIndex        =   28
      Top             =   6180
      Width           =   915
   End
   Begin VB.Line Line2 
      X1              =   8700
      X2              =   8700
      Y1              =   0
      Y2              =   6600
   End
   Begin VB.Label Label7 
      Caption         =   "Origin:"
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label Label6 
      Caption         =   "Points"
      Height          =   255
      Left            =   7260
      TabIndex        =   16
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "X                     Y                     Z"
      Height          =   255
      Left            =   5700
      TabIndex        =   14
      Top             =   1980
      Width           =   2835
   End
   Begin VB.Label lblSide 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   6420
      TabIndex        =   9
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Editing Side:"
      Height          =   195
      Left            =   5040
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "No Sides:"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "VB Name:"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   4980
      X2              =   8520
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Properties:"
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   60
      Width           =   3615
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private polybase As vx_Polygon
Private polySides(0 To 50) As vx_PolySide

Private vx_fps2 As Long

Private Sub Command1_Click()
If lblSide = 0 Then Exit Sub

SaveToVar lblSide

lblSide = lblSide - 1

LoadFromVar lblSide

End Sub

Private Sub Command2_Click()

If lblSide = txtNoSides - 1 Then Exit Sub

SaveToVar lblSide

lblSide = lblSide + 1

LoadFromVar lblSide

End Sub

Private Sub Command3_Click()

SaveToVar lblSide
DoEvents

Picture1.Cls

DrawPolygon polybase, polySides

Picture1.Refresh
DoEvents

End Sub

Private Sub Command4_Click()

If Timer1.Enabled = True Then

    Timer1.Enabled = False

Else

    Timer1.Enabled = True
    
End If

End Sub

Private Sub Command5_Click()

SaveToVar lblSide

Code = ""

Code = Code & "'Auto Generated Using The Visual X Editor" & vbCrLf

Code = Code & "Dim " & txtVBName & " as vx_Polygon" & vbCrLf
Code = Code & "Dim " & txtVBName & "Sides(" & polybase.sides - 1 & ") as vx_PolySide" & vbCrLf

Code = Code & txtVBName & ".sides = " & polybase.sides & vbCrLf
Code = Code & txtVBName & ".origin.x = " & polybase.origin.X & vbCrLf
Code = Code & txtVBName & ".origin.y = " & polybase.origin.Y & vbCrLf
Code = Code & txtVBName & ".origin.z = " & polybase.origin.Z & vbCrLf

Dim i As Long, o As Long

For i = 0 To polybase.sides - 1

    Code = Code & txtVBName & "Sides(" & i & ").c_points = " & polySides(i).c_points & vbCrLf
    
    o = 0
    
    For o = 0 To polySides(i).c_points - 1
    
        Code = Code & txtVBName & "Sides(" & i & ").points(" & o & ").x = " & polySides(i).points(o).X & vbCrLf
        Code = Code & txtVBName & "Sides(" & i & ").points(" & o & ").y = " & polySides(i).points(o).Y & vbCrLf
        Code = Code & txtVBName & "Sides(" & i & ").points(" & o & ").z = " & polySides(i).points(o).Z & vbCrLf

    Next o
    
Next i

End Sub

Private Sub Form_Load()
Dim i As Long

Do Until i = 51

    polySides(i).c_points = 1
    
    i = i + 1
    
Loop

modVisualX.Canvas = Picture1

End Sub

Private Sub SX_Change()
SaveToVar lblSide
DoEvents

polybase.angle.X = SX

Picture1.Cls

DrawPolygon polybase, polySides

Picture1.Refresh
DoEvents

End Sub

Private Sub Sy_Change()
SaveToVar lblSide
DoEvents

polybase.angle.Y = SY

Picture1.Cls

DrawPolygon polybase, polySides

Picture1.Refresh
DoEvents

End Sub

Private Sub Sz_Change()
SaveToVar lblSide
DoEvents

polybase.angle.Z = SZ

Picture1.Cls

DrawPolygon polybase, polySides

Picture1.Refresh
DoEvents

End Sub

Private Sub Timer1_Timer()

SaveToVar lblSide
DoEvents

polybase.angle.X = polybase.angle.X + 1
polybase.angle.Y = polybase.angle.Y + 1
polybase.angle.Z = polybase.angle.Z + 1

If polybase.angle.X = 360 Then polybase.angle.X = 0
If polybase.angle.Y = 360 Then polybase.angle.Y = 0
If polybase.angle.Z = 360 Then polybase.angle.Z = 0

Picture1.Cls

DrawPolygon polybase, polySides

vx_fps2 = vx_fps2 + 1

Picture1.Refresh
DoEvents

End Sub

Private Sub Timer2_Timer()
lblFPS = vx_fps2
vx_fps2 = 0
End Sub

Private Sub txtP_Change()

If txtP = "" Or txtP = "0" Then Exit Sub

Dim i As Long

i = txtP - 1

If i = txtX.UBound Then Exit Sub

If i > txtX.UBound Then

    Do Until txtX.UBound = i

        Load txtX(txtX.UBound + 1)
        DoEvents
        
        txtX(txtX.UBound).Top = txtX(txtX.UBound - 1).Top + txtX(txtX.UBound - 1).Height + 60
        
        txtX(txtX.UBound).Visible = True
    
        Load txtY(txtY.UBound + 1)
        DoEvents
        
        txtY(txtY.UBound).Top = txtX(txtX.UBound - 1).Top + txtY(txtY.UBound - 1).Height + 60
    
        txtY(txtY.UBound).Visible = True
    
        Load txtZ(txtZ.UBound + 1)
        DoEvents
        
        txtZ(txtZ.UBound).Top = txtX(txtX.UBound - 1).Top + txtZ(txtZ.UBound - 1).Height + 60
        
        txtZ(txtZ.UBound).Visible = True
    
    Loop
    
Else

    Do Until txtX.UBound = i

        Unload txtX(txtX.UBound)
    
        Unload txtY(txtY.UBound)
    
        Unload txtZ(txtZ.UBound)
        
    Loop

End If
    
End Sub

Public Function SaveToVar(sideNo)

On Error Resume Next
If txtP = "" Then Exit Function

Dim i As Long

polybase.origin.X = txtOX
polybase.origin.Y = txtOY
polybase.origin.Z = txtOZ

polybase.sides = txtNoSides

polySides(sideNo).c_points = txtP
Do Until i = txtP

    If txtX(i) <> "" And IsNumeric(txtX(i)) Then polySides(sideNo).points(i).X = txtX(i)
    If txtY(i) <> "" And IsNumeric(txtY(i)) Then polySides(sideNo).points(i).Y = txtY(i)
    If txtZ(i) <> "" And IsNumeric(txtZ(i)) Then polySides(sideNo).points(i).Z = txtZ(i)
    
    i = i + 1

Loop
End Function

Public Function LoadFromVar(sideNo)

Dim i As Long

txtOX = polybase.origin.X
txtOY = polybase.origin.Y
txtOZ = polybase.origin.Z

txtNoSides = polybase.sides

txtP = polySides(sideNo).c_points
DoEvents

Do Until i = txtP

    txtX(i) = polySides(sideNo).points(i).X
    txtY(i) = polySides(sideNo).points(i).Y
    txtZ(i) = polySides(sideNo).points(i).Z
    
    i = i + 1

Loop

End Function
