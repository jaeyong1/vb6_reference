Attribute VB_Name = "Module1"
Option Explicit

Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const RGN_DIFF = 4
Public Const RGN_COPY = 5
Public Const RGN_MAX = RGN_COPY
Public Const RGN_MIN = RGN_AND

Public Const GWL_EXSTYLE = (-20)
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const WS_EX_LAYERED = &H80000

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function SetWindowRgn Lib "user32" _
      (ByVal hWnd As Long, _
       ByVal hRgn As Long, _
       ByVal bRedraw As Boolean) As Long
       
Public Declare Function CreateRectRgn Lib "gdi32" _
      (ByVal X1 As Long, _
       ByVal Y1 As Long, _
       ByVal X2 As Long, _
       ByVal Y2 As Long) As Long
       
Public Declare Function CreateEllipticRgn Lib "gdi32" _
      (ByVal X1 As Long, _
       ByVal Y1 As Long, _
       ByVal X2 As Long, _
       ByVal Y2 As Long) As Long
       
Public Declare Function CreateRoundRectRgn Lib "gdi32" _
      (ByVal X1 As Long, _
       ByVal Y1 As Long, _
       ByVal X2 As Long, _
       ByVal Y2 As Long, _
       ByVal X3 As Long, _
       ByVal Y3 As Long) As Long

Public Declare Function CreatePolygonRgn Lib "gdi32" _
      (lpPoint As POINTAPI, _
       ByVal nCount As Long, _
       ByVal nPolyFillMode As Long) As Long

Public Declare Function CombineRgn Lib "gdi32" _
      (ByVal hDestRgn As Long, _
       ByVal hSrcRgn1 As Long, _
       ByVal hSrcRgn2 As Long, _
       ByVal nCombineMode As Long) As Long
       
Public Sub MakeShowOnlyControl(paramForm As Form, paramHWND As Long)
    Dim REGN As Long
    Dim TmpREGN As Long
    Dim Control As Control
    Dim LinePoints(4) As POINTAPI
    
    REGN = CreateRectRgn(0, 0, 0, 0)
    
    For Each Control In paramForm.Controls
        'If the control is a line...
        If TypeOf Control Is Line Then
            'Checks the slope
            If Abs((Control.Y1 - Control.Y2) / (Control.X1 - Control.X2)) > 1 Then
                'If it's more verticle than horizontal then
                'Set the points
                LinePoints(0).X = Control.X1 - 1
                LinePoints(0).Y = Control.Y1
                LinePoints(1).X = Control.X2 - 1
                LinePoints(1).Y = Control.Y2
                LinePoints(2).X = Control.X2 + 1
                LinePoints(2).Y = Control.Y2
                LinePoints(3).X = Control.X1 + 1
                LinePoints(3).Y = Control.Y1
            Else
                'If it's more horizontal than verticle then
                'Set the points
                LinePoints(0).X = Control.X1
                LinePoints(0).Y = Control.Y1 - 1
                LinePoints(1).X = Control.X2
                LinePoints(1).Y = Control.Y2 - 1
                LinePoints(2).X = Control.X2
                LinePoints(2).Y = Control.Y2 + 1
                LinePoints(3).X = Control.X1
                LinePoints(3).Y = Control.Y1 + 1
            End If
            'Creates the new polygon with the points
            TmpREGN = CreatePolygonRgn(LinePoints(0), 4, 1)
            
        'If the control is a shape...
        ElseIf TypeOf Control Is Shape Then
            
            'An if that checks the type
            If Control.Shape = 0 Then
            'It's a rectangle
                TmpREGN = CreateRectRgn(Control.Left, Control.Top, Control.Left + Control.Width, Control.Top + Control.Height)
            ElseIf Control.Shape = 1 Then
            'It's a square
                If Control.Width < Control.Height Then
                    TmpREGN = CreateRectRgn(Control.Left, Control.Top + (Control.Height - Control.Width) / 2, Control.Left + Control.Width, Control.Top + (Control.Height - Control.Width) / 2 + Control.Width)
                Else
                    TmpREGN = CreateRectRgn(Control.Left + (Control.Width - Control.Height) / 2, Control.Top, Control.Left + (Control.Width - Control.Height) / 2 + Control.Height, Control.Top + Control.Height)
                End If
            ElseIf Control.Shape = 2 Then
            'It's an oval
                TmpREGN = CreateEllipticRgn(Control.Left, Control.Top, Control.Left + Control.Width + 0.5, Control.Top + Control.Height + 0.5)
            ElseIf Control.Shape = 3 Then
            'It's a circle
                If Control.Width < Control.Height Then
                    TmpREGN = CreateEllipticRgn(Control.Left, Control.Top + (Control.Height - Control.Width) / 2, Control.Left + Control.Width + 0.5, Control.Top + (Control.Height - Control.Width) / 2 + Control.Width + 0.5)
                Else
                    TmpREGN = CreateEllipticRgn(Control.Left + (Control.Width - Control.Height) / 2, Control.Top, Control.Left + (Control.Width - Control.Height) / 2 + Control.Height + 0.5, Control.Top + Control.Height + 0.5)
                End If
            ElseIf Control.Shape = 4 Then
            'It's a rounded rectangle
                If Control.Width > Control.Height Then
                    TmpREGN = CreateRoundRectRgn(Control.Left, Control.Top, Control.Left + Control.Width + 1, Control.Top + Control.Height + 1, Control.Height / 4, Control.Height / 4)
                Else
                    TmpREGN = CreateRoundRectRgn(Control.Left, Control.Top, Control.Left + Control.Width + 1, Control.Top + Control.Height + 1, Control.Width / 4, Control.Width / 4)
                End If
            ElseIf Control.Shape = 5 Then
            'It's a rounded square
                If Control.Width > Control.Height Then
                    TmpREGN = CreateRoundRectRgn(Control.Left + (Control.Width - Control.Height) / 2, Control.Top, Control.Left + (Control.Width - Control.Height) / 2 + Control.Height + 1, Control.Top + Control.Height + 1, Control.Height / 4, Control.Height / 4)
                Else
                    TmpREGN = CreateRoundRectRgn(Control.Left, Control.Top + (Control.Height - Control.Width) / 2, Control.Left + Control.Width + 1, Control.Top + (Control.Height - Control.Width) / 2 + Control.Width + 1, Control.Width / 4, Control.Width / 4)
                End If
            End If
            
            'If the control is a shape with a transparent background
            If Control.BackStyle = 0 Then
                
                'Combines the regions in memory and makes a new one
                CombineRgn REGN, REGN, TmpREGN, RGN_XOR
                
                If Control.Shape = 0 Then
                'Rectangle
                    TmpREGN = CreateRectRgn(Control.Left + 1, Control.Top + 1, Control.Left + Control.Width - 1, Control.Top + Control.Height - 1)
                ElseIf Control.Shape = 1 Then
                'Square
                    If Control.Width < Control.Height Then
                        TmpREGN = CreateRectRgn(Control.Left + 1, Control.Top + (Control.Height - Control.Width) / 2 + 1, Control.Left + Control.Width - 1, Control.Top + (Control.Height - Control.Width) / 2 + Control.Width - 1)
                    Else
                        TmpREGN = CreateRectRgn(Control.Left + (Control.Width - Control.Height) / 2 + 1, Control.Top + 1, Control.Left + (Control.Width - Control.Height) / 2 + Control.Height - 1, Control.Top + Control.Height - 1)
                    End If
                ElseIf Control.Shape = 2 Then
                'Oval
                    TmpREGN = CreateEllipticRgn(Control.Left + 1, Control.Top + 1, Control.Left + Control.Width - 0.5, Control.Top + Control.Height - 0.5)
                ElseIf Control.Shape = 3 Then
                'Circle
                    If Control.Width < Control.Height Then
                        TmpREGN = CreateEllipticRgn(Control.Left + 1, Control.Top + (Control.Height - Control.Width) / 2 + 1, Control.Left + Control.Width - 0.5, Control.Top + (Control.Height - Control.Width) / 2 + Control.Width - 0.5)
                    Else
                        TmpREGN = CreateEllipticRgn(Control.Left + (Control.Width - Control.Height) / 2 + 1, Control.Top + 1, Control.Left + (Control.Width - Control.Height) / 2 + Control.Height - 0.5, Control.Top + Control.Height - 0.5)
                    End If
                ElseIf Control.Shape = 4 Then
                'Rounded rectangle
                    If Control.Width > Control.Height Then
                        TmpREGN = CreateRoundRectRgn(Control.Left + 1, Control.Top + 1, Control.Left + Control.Width, Control.Top + Control.Height, Control.Height / 4, Control.Height / 4)
                    Else
                        TmpREGN = CreateRoundRectRgn(Control.Left + 1, Control.Top + 1, Control.Left + Control.Width, Control.Top + Control.Height, Control.Width / 4, Control.Width / 4)
                    End If
                ElseIf Control.Shape = 5 Then
                'Rounded square
                    If Control.Width > Control.Height Then
                        TmpREGN = CreateRoundRectRgn(Control.Left + (Control.Width - Control.Height) / 2 + 1, Control.Top + 1, Control.Left + (Control.Width - Control.Height) / 2 + Control.Height, Control.Top + Control.Height, Control.Height / 4, Control.Height / 4)
                    Else
                        TmpREGN = CreateRoundRectRgn(Control.Left + 1, Control.Top + (Control.Height - Control.Width) / 2 + 1, Control.Left + Control.Width, Control.Top + (Control.Height - Control.Width) / 2 + Control.Width, Control.Width / 4, Control.Width / 4)
                    End If
                End If
            End If
        Else
            'Create a rectangular region with its parameters
            TmpREGN = CreateRectRgn(Control.Left, Control.Top, Control.Left + Control.Width, Control.Top + Control.Height)
        End If
        
        CombineRgn REGN, REGN, TmpREGN, RGN_XOR
    Next
    
    SetWindowRgn paramHWND, REGN, True
End Sub
