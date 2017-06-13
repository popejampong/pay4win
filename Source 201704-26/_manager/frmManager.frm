VERSION 5.00
Begin VB.Form frmManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Authorization Required"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1350
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   555
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2160
      TabIndex        =   3
      Top             =   1050
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   540
      TabIndex        =   2
      Top             =   1050
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1350
      TabIndex        =   0
      Top             =   165
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   165
      TabIndex        =   5
      Top             =   570
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Manager ID:"
      Height          =   270
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "frmManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Cost Accounting System
' module        :   frmManager
' programmer    :   _-=[ srm ]=-_
' date          :   7 feb 2005

Option Explicit
    
Private Sub cmdCancel_Click()
    ModalResult = mrCancel
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim cSqlStmt As String
    cSqlStmt = "SELECT * FROM PA2360 WHERE (USERLEVEL=1) " & _
               IIf(Val(Tag) = 1, " and (groupid>0)", "") & _
               " AND (USERID=" & cQuote & txtUserName.Text & cQuote & ")" & _
               " AND (AES_DECRYPT(PASSWORD,UCASE(USERID))=" & cQuote & txtPassword.Text & cQuote & ")"
    OpenQueryDNS cSqlStmt, objdbRs, False
    ModalResult = IIf(Not objdbRs.EOF, mrOk, mrNone)
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Form_Load()
    OpenQueryDNS "SELECT * FROM PA2360 WHERE USERLEVEL=1 ORDER BY USERID", objdbRs, False
    txtUserName.MaxLength = objdbRs.Fields.Item("USERID").DefinedSize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
