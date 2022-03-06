VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Smallest Chart (VB6)"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18630
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1242
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox graf_val 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   240
      ScaleHeight     =   415
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1207
      TabIndex        =   0
      Top             =   240
      Width           =   18135
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


Function chart(g, c, e)

    sig = Split(g, ",")
    
    mx = 0
    mn = 0
    
    For i = 0 To UBound(sig)
        If (Val(sig(i)) > mx) Then mx = Val(sig(i))
        If (Val(sig(i)) < mn) Then mn = Val(sig(i))
    Next i

    w = graf_val.ScaleWidth
    h = graf_val.ScaleHeight

    d = (w - 80) / (UBound(sig) - 1)
    
    If (e = "|") Then
        graf_val.Cls
        mxg = mx
        mng = mn
    End If
    
    graf_val.DrawWidth = 4
    
    For i = 0 To UBound(sig) - 1
    
        y = h - 15 - ((h - 15) / (mx - mn)) * (Val(sig(i)) - mn)
        x = d * i

        If (i = 0) Then
            oldX = x
            oldY = y
        End If
        
        graf_val.Line (oldX, oldY)-(x, y), c
        
        oldX = x
        oldY = y
        
    Next i
 
End Function


Private Sub Form_Load()
    
    Form1.DrawWidth = 3
    
    A = "0,2.62,5.23,7.82,10.4,12.94,15.45,17.92,20.34,22.7,25,27.23,29.39,31.47,33.46,35.36,37.16,38.86,40.45,41.93,43.3,44.55,45.68,46.68,47.55,48.3,48.91,49.38,49.73,49.93,50,49.93,49.73,49.38,48.91,48.3,47.55,46.68,45.68,44.55,43.3,41.93,40.45,38.86,37.16,35.36,33.46,31.47,29.39,27.23,25,22.7,20.34,17.92,15.45,12.94,10.4,7.82,5.23,2.62,0"
    B = "0,0.14,0.29,0.45,0.64,0.86,1.14,1.53,2.13,3.27,6.41,75.31,7.75,3.61,2.29,1.62,1.2,0.9,0.67,0.48,0.32,0.17,0.03,0.12,0.26,0.42,0.6,0.81,1.08,1.44,2,2.99,5.45,25.09,9.79,4.03,2.47,1.72,1.27,0.95,0.71,0.52,0.35,0.2,0.05,0.09,0.23,0.39,0.56,0.77,1.02,1.36,1.87,2.74,4.74,15.04,13.27,4.54,2.67,1.83,1.34"

    Call chart(B, vbRed, "|")
  
End Sub
