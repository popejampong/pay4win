VERSION 5.00
Begin VB.Form frmSpecialAssess 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Special Assessment Amount"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAmount 
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      Top             =   345
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   705
      TabIndex        =   2
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1905
      TabIndex        =   3
      Top             =   1020
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "Enter Amount"
      Height          =   270
      Index           =   0
      Left            =   495
      TabIndex        =   0
      Top             =   360
      Width           =   1080
   End
End
Attribute VB_Name = "frmSpecialAssess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmSpecialAssess
' description   :   module for Special Assessment amount...
' programmer    :   _-=[ srm ]=-_
' date          :   25 may 2006

Option Explicit

Private Sub cmdCancel_Click()
    ModalResult = mrCancel
    Unload Me
End Sub

Private Sub cmdOK_Click()
    nAssessAmt = Val(txtAmount.Text)
    ModalResult = mrOk
    Unload Me
End Sub

Private Sub txtAmount_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtAmount.Text)
End Sub
