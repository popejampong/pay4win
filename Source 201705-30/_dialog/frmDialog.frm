VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm Exit..."
   ClientHeight    =   1140
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Select Option:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmDialog
' description   :   messagebox to log-off/exit application
' programmer    :   _-=[ srm ]=-_
' date          :   17 Oct 2005

Option Explicit
Dim oTempADO As New ADODB.Recordset

Private Sub CancelButton_Click()
    nAccess_Tag = Combo1.ListIndex
    ModalResult = mrCancel
    Unload Me
End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Log Off " & Trim(gUserName)
    Combo1.AddItem "Exit Application"
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If Combo1.ListIndex = 0 Then
'        frmMain.showLogin
'    Else
'        ' --> set status of current user to 0 02072005
'        OpenQueryDNS "UPDATE DI2360 SET STATUS=0 WHERE USERID=" & cQuote & gUserID & cQuote, objdbRs, True
'        Script2File "UPDATE DI2360 SET STATUS=0 WHERE USERID=" & cQuote & gUserID & cQuote
'        Log2Audit "frmMain", "User logged-off."
'        Write2File gUserID & " logged-off at " & Now
'        Write2File ""
'
'        End
'    End If
End Sub

Private Sub OKButton_Click()
    nAccess_Tag = Combo1.ListIndex
    ModalResult = mrOk
    Unload Me
End Sub
