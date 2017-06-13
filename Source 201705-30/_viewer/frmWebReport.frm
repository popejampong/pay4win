VERSION 5.00
Begin VB.Form frmWebReport 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox WebBrowser1 
      Height          =   4935
      Left            =   105
      ScaleHeight     =   4875
      ScaleWidth      =   6600
      TabIndex        =   0
      Top             =   360
      Width           =   6660
   End
End
Attribute VB_Name = "frmWebReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    With WebBrowser1
        .top = ScaleTop
        .left = ScaleLeft
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub
