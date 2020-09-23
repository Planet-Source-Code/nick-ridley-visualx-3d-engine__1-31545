Attribute VB_Name = "modVisualX"
'''''''''''''''''''''''''''''
'***************************'
'*      V I S U A L X      *'
'***************************'
'''''''''''''''''''''''''''''
'CopyRight © 2001 Spyder-Net'
'''''''''''''''''''''''''''''

'VisualX Object Types

Public Type vx_Single
    X As Long
    Y As Long
    Z As Long
End Type

Public Type vx_Angle
    X As Double
    Y As Double
    Z As Double
End Type

Public Type vx_Polygon
    sides As Long
    origin As vx_Single
    angle As vx_Angle
    draworder(1000) As Long
End Type

Public Type vx_PolySide
    c_points As Long
    points(0 To 25) As vx_Single
    tp(0 To 25) As vx_Single
    normal As vx_Single
    fBrush As Long
    az As Double
    order As Long
End Type

'VisualX System Variables
Private vx_Canvas As Object
Private vx_CanvasBG As Long
'Private vx_ClearPen

Private CosA(0 To 359) As Double
Private SinA(0 To 359) As Double

'VisualX Utilised WinAPI

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long



Public Property Get Canvas() As Object

Set Canvas = vx_Canvas

End Property

Public Property Let Canvas(ByVal val As Object)

Set vx_Canvas = val

vx_Canvas.BackColor = RGB(20, 80, 200)
vx_CanvasBG = RGB(20, 80, 200)

vx_Canvas.ScaleMode = vbPixels

End Property

Public Function InitVisualX()

Dim i As Double

For i = o To 359

    CosA(i) = Cos(i * (3.14159265358979 / 180))
    SinA(i) = Sin(i * (3.14159265358979 / 180))

Next i

End Function


Public Function DrawPolygon(poly As vx_Polygon, sides() As vx_PolySide)

Dim old As vx_Single
Dim i As Long, o As Long
Dim sp(32768) As POINTAPI

'calculate angle changes

For i = 0 To poly.sides - 1

    o = 0
    
    'FindNormals sides, poly.sides
    
    For o = 0 To sides(i).c_points - 1
    
        With sides(i).points(o)
    
            sides(i).tp(o) = sides(i).points(o)
    
            'x change
            old = sides(i).tp(o)
            
            sides(i).tp(o).X = old.X
            sides(i).tp(o).Y = old.Y * CosA(poly.angle.X) - old.Z * SinA(poly.angle.X)
            sides(i).tp(o).Z = old.Z * CosA(poly.angle.X) + old.Y * SinA(poly.angle.X)
            
            'y change
            old = sides(i).tp(o)
            
            sides(i).tp(o).Y = old.Y
            sides(i).tp(o).Z = old.Z * CosA(poly.angle.Y) - old.X * SinA(poly.angle.Y)
            sides(i).tp(o).X = old.X * CosA(poly.angle.Y) + old.Z * SinA(poly.angle.Y)
            
            'z change
            old = sides(i).tp(o)
            
            sides(i).tp(o).Z = old.Z
            sides(i).tp(o).Y = old.Y * CosA(poly.angle.Z) - old.X * SinA(poly.angle.Z)
            sides(i).tp(o).X = old.X * CosA(poly.angle.Z) + old.Y * SinA(poly.angle.Z)
            
            sides(i).tp(o).X = sides(i).tp(o).X + poly.origin.X
            sides(i).tp(o).Y = sides(i).tp(o).Y + poly.origin.Y
            sides(i).tp(o).Z = sides(i).tp(o).Z + poly.origin.Z
            
        End With
        
    Next o
    
Next i

i = 0
o = 0

'if you can get this to work leave me a message
'CalcOrder sides, poly, poly.sides
'DoEvents

'calculate and draw sides
'For i = 0 To poly.sides - 1

'    o = 0
    
'    For o = 0 To sides(poly.draworder(i)).c_points - 1
    
'        With sides(poly.draworder(i)).tp(o)
    
'            sp(o).X = .X
'            sp(o).Y = .Y
            
'        End With
        
'    Next o
    
    'If sides(poly.draworder(i)).fBrush <> 0 Then SelectObject vx_Canvas.hdc, sides(poly.draworder(i)).fBrush
    
    Polygon vx_Canvas.hdc, sp(0), sides(poly.draworder(i)).c_points
        
For i = 0 To poly.sides - 1

    o = 0
    
    For o = 0 To sides(i).c_points - 1
    
        With sides(i).tp(o)
    
            sp(o).X = .X
            sp(o).Y = .Y
            
        End With
        
    Next o
    
    
    Polygon vx_Canvas.hdc, sp(0), sides(i).c_points
        
        
Next i

End Function

Private Function FindNormals(sides() As vx_PolySide, num As Long)
    'This function finds the normal of each plane
    
    For i = 0 To num - 1
        'Find the normal vector
        With sides(i)
            '    *                           *   *                                    *   *                           *   *                                    *
            Nx = (.points(1).Y - .points(0).Y) * (.points(.c_points).Z - .points(0).Z) - (.points(1).Z - .points(0).Z) * (.points(.c_points).Y - .points(0).Y)
            Ny = (.points(1).Z - .points(0).Z) * (.points(.c_points).X - .points(0).X) - (.points(1).X - .points(0).X) * (.points(.c_points).Z - .points(0).Z)
            Nz = (.points(1).X - .points(0).X) * (.points(.c_points).Y - .points(0).Y) - (.points(1).Y - .points(0).Y) * (.points(.c_points).X - .points(0).X)
        
        
            'Normalize the normal vector (make length of 1)
            length = Sqr(Nx ^ 2 + Ny ^ 2 + Nz ^ 2)
            
            If length <> 0 Then
            .normal.X = Nx / length
            .normal.Y = Ny / length
            .normal.Z = Nz / length
            End If
        End With
    Next i
    
End Function

Private Function VisiblePlane(shape As vx_PolySide, CameraX As Integer, CameraY As Integer, CameraZ As Integer)
    'this function takes the normal of the plane and returns True if visible FALSE if
    'not visible
    
    'Camera is the spot the object is being viewed from
    
    'Find the dot product
    D = (shape.normal.X * CameraX) + (shape.normal.Y * CameraY) + (shape.normal.Z * CameraZ)
    
    'return true if object is visible
    VisiblePlane = D >= 0
    
End Function

'This dont work properly

Public Function CalcOrder(sides() As vx_PolySide, poly As vx_Polygon, sc As Long)

'*************************************************************
'*IMPORTANT: i have no idea how to get this working properly!*
'*************************************************************

Dim Aa As Long, Ab As Long
Dim i As Long
        
'to be honest no polygon would ever have this many sides, but hell
'anything is possible
Dim td(1000) As Double
    
For Aa = 0 To sc - 1

    i = 0
    
    For i = 0 To sides(Aa).c_points - 1
    
        sides(Aa).az = sides(Aa).az + sides(Aa).tp(i).Z
    
    Next i
    
    sides(Aa).az = sides(Aa).az / sides(Aa).c_points

Next Aa
        
Aa = 1
        
For Aa = 1 To sc - 1
    
    Ab = Aa
    
ca:
    
    If sides(Aa).az > sides(td(Ab - 1)).az Then
    
        td(Ab) = td(Ab - 1)
        td(Ab - 1) = Aa
    
    Ab = Ab - 1
    
    If Ab > 0 Then GoTo ca
    
    Else
    
    td(Aa) = Aa
    
    End If
    
Next Aa

Aa = 0
        
For Aa = 0 To sc - 1
    
    poly.draworder(Aa) = td(Aa)
    
Next Aa

End Function
