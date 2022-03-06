VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Shortest Chart with axes (VB6)"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20595
   LinkTopic       =   "Form1"
   ScaleHeight     =   536
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1373
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VScroll2 
      Height          =   2895
      Left            =   19800
      Max             =   300
      TabIndex        =   4
      Top             =   3480
      Value           =   50
      Width           =   495
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2895
      Left            =   19800
      Max             =   300
      TabIndex        =   3
      Top             =   240
      Value           =   150
      Width           =   495
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   495
      Left            =   360
      Max             =   300
      TabIndex        =   2
      Top             =   6600
      Value           =   50
      Width           =   9255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   10320
      Max             =   300
      TabIndex        =   1
      Top             =   6600
      Value           =   150
      Width           =   9255
   End
   Begin VB.PictureBox graf_val 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   360
      ScaleHeight     =   407
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1279
      TabIndex        =   0
      Top             =   240
      Width           =   19215
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   -4680
      Picture         =   "chart.frx":0000
      Top             =   7320
      Width           =   25290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'############################################################################################################################## -->
'#  John Wiley & Sons, Inc.                                                                                                   # -->
'#                                                                                                                            # -->
'#  Book:   Algorithms in Bioinformatics: Theory and Implementation                                                           # -->
'#  Author: Dr. Paul A. Gagniuc                                                                                               # -->
'#                                                                                                                            # -->
'#  Institution:                                                                                                              # -->
'#    University Politehnica of Bucharest                                                                                     # -->
'#    Faculty of Engineering in Foreign Languages                                                                             # -->
'#    Department of Engineering in Foreign Languages                                                                          # -->
'#                                                                                                                            # -->
'#  Area:   European Union                                                                                                    # -->
'#  Date:   04/01/2021                                                                                                        # -->
'#                                                                                                                            # -->
'#  Mode:   Visual Basic 6.0                                                                                                  # -->
'#                                                                                                                            # -->
'#  Cite this work as:                                                                                                        # -->
'#    Paul A. Gagniuc. Algorithms in Bioinformatics: Theory and Implementation. John Wiley & Sons, 2021, ISBN: 9781119697961. # -->
'#                                                                                                                            # -->
'############################################################################################################################## -->

Dim A As String
Dim B As String
Dim mxg As Variant
Dim mng As Variant
Dim mxAB As Variant

Dim wO As Integer
Dim hO As Integer
Dim wV As Integer
Dim hV As Integer

Function chart(g, c, e)

    sig = Split(g, ",")
    
    mx = 0
    mn = 0
    
    For i = 0 To UBound(sig)
        If (Val(sig(i)) > mx) Then mx = Val(sig(i))
        If (Val(sig(i)) < mn) Then mn = Val(sig(i))
    Next i

    w = graf_val.ScaleWidth - wO
    h = graf_val.ScaleHeight - hO

    d = (w - 80) / (UBound(sig) - 1)
    
    If (e = "|") Then
    
        graf_val.Cls
        mxg = mx
        mng = mn
        
    End If
    
    graf_val.DrawWidth = 4
    
    For i = 0 To UBound(sig) - 1
    
        y = hV + (h - 15 - ((h - 15) / (mx - mn)) * (Val(sig(i)) - mn))
        x = wV + (d * i)

        If (i = 0) Then
            oldX = x
            oldY = y
        End If
        
        graf_val.Line (oldX, oldY)-(x, y), c
        
        oldX = x
        oldY = y
        
    Next i
 
    Call draw_scale(UBound(sig) - 1, wO, hO, wV, hV)
 
End Function


Private Sub Form_Load()
    
    Form1.DrawWidth = 3
    
    A = "0,2.62,5.23,7.82,10.4,12.94,15.45,17.92,20.34,22.7,25,27.23,29.39,31.47,33.46,35.36,37.16,38.86,40.45,41.93,43.3,44.55,45.68,46.68,47.55,48.3,48.91,49.38,49.73,49.93,50,49.93,49.73,49.38,48.91,48.3,47.55,46.68,45.68,44.55,43.3,41.93,40.45,38.86,37.16,35.36,33.46,31.47,29.39,27.23,25,22.7,20.34,17.92,15.45,12.94,10.4,7.82,5.23,2.62,0"
    B = "0,0.14,0.29,0.45,0.64,0.86,1.14,1.53,2.13,3.27,6.41,75.31,7.75,3.61,2.29,1.62,1.2,0.9,0.67,0.48,0.32,0.17,0.03,0.12,0.26,0.42,0.6,0.81,1.08,1.44,2,2.99,5.45,25.09,9.79,4.03,2.47,1.72,1.27,0.95,0.71,0.52,0.35,0.2,0.05,0.09,0.23,0.39,0.56,0.77,1.02,1.36,1.87,2.74,4.74,15.04,13.27,4.54,2.67,1.83,1.34"


    wO = HScroll1.Value
    wV = HScroll2.Value
    hO = VScroll1.Value
    hV = VScroll2.Value

    Call chart(B, vbRed, "|")
    
End Sub


Function draw_scale(ByVal k_stat As Integer, wO, hO, wV, hV)

    Dim zx, qx, zy, qy As Variant
    Dim sp As Variant
    Dim i As Integer

    'X axis on graf_val OBJ
    '-------------------------------------
    sp = ((graf_val.Width - wO) / k_stat)
    
    For i = 0 To k_stat
    
        zx = wV + (sp * i)
        qx = zx
        zy = graf_val.Height - hO + hV
        qy = graf_val.Height - hO + hV + 10
    
        If k_stat < 100 Then
            graf_val.CurrentX = zx - 8
            graf_val.CurrentY = qy + 5
            graf_val.Print i
        End If
    
        graf_val.Line (zx, zy)-(qx, qy), &H808080
    
        If (i = k_stat Or i = 0) Then
            graf_val.Line (zx, zy)-(qx, hV), &H808080
        End If
        
    Next i
    '-------------------------------------

    'Y axis on graf_val OBJ
    '-------------------------------------
    zx = wV
    qx = wV + graf_val.Width - wO + 10
    zy = hV
    qy = zy
    graf_val.Line (zx, zy)-(qx, qy), &H808080
    
    graf_val.CurrentX = qx + 2
    graf_val.CurrentY = qy - 6
    graf_val.Print mxg

    '-------------------------------------
    
    zx = wV
    qx = wV + graf_val.Width - wO + 10
    zy = hV + graf_val.Height - hO
    qy = zy
    graf_val.Line (zx, zy)-(qx, qy), &H808080
    
    graf_val.CurrentX = qx + 2
    graf_val.CurrentY = qy - 6
    graf_val.Print mng
    '-------------------------------------

End Function


Private Sub HScroll1_Scroll()
    wO = HScroll1.Value
    Call chart(B, vbRed, "|")
End Sub

Private Sub HScroll2_Scroll()
    wV = HScroll2.Value
    Call chart(B, vbRed, "|")
End Sub

Private Sub VScroll1_Scroll()
    hO = VScroll1.Value
    Call chart(B, vbRed, "|")
End Sub

Private Sub VScroll2_Scroll()
    hV = VScroll2.Value
    Call chart(B, vbRed, "|")
End Sub
