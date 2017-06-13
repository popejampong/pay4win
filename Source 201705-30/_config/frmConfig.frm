VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ODBC Configuration Entry"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   9105
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "TXT:ODBCDATABASE"
      Top             =   1905
      Width           =   3180
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "TXT:ODBCUSER"
      Top             =   1305
      Width           =   3180
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "TXT:ODBCPASSWORD"
      Top             =   1605
      Width           =   3180
   End
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   15
      TabIndex        =   20
      Top             =   2205
      Width           =   9075
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7095
         Picture         =   "frmConfig.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "20"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6255
         Picture         =   "frmConfig.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "19"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5415
         Picture         =   "frmConfig.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "18"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmConfig.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "12"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmConfig.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "14"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmConfig.frx":8662
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "13"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8070
         Picture         =   "frmConfig.frx":9FE4
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "21"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4575
         Picture         =   "frmConfig.frx":B966
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "17"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3600
         Picture         =   "frmConfig.frx":D2E8
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "15"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmConfig.frx":EC6A
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "11"
         Top             =   165
         Width           =   855
      End
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "TXT:ODBCSERVER"
      Top             =   1005
      Width           =   3180
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:DESCRIPTION"
      Top             =   705
      Width           =   3180
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "TXT:ODBCCODE"
      Top             =   105
      Width           =   960
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "TXT:DSOURCENAME"
      Top             =   405
      Width           =   3180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DATABASE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   105
      TabIndex        =   24
      Top             =   1935
      Width           =   1680
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "USER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   105
      TabIndex        =   23
      Top             =   1350
      Width           =   1680
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SERVER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   105
      TabIndex        =   22
      Top             =   1035
      Width           =   1680
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   105
      TabIndex        =   21
      Top             =   1635
      Width           =   1680
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPTION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   105
      TabIndex        =   19
      Top             =   735
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   100
      TabIndex        =   18
      Top             =   165
      Width           =   1680
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DATA SOURCE NAME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   105
      TabIndex        =   17
      Top             =   450
      Width           =   1680
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   4515
      Left            =   0
      Top             =   0
      Width           =   1905
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Cost Accounting System
' module        :   frmComfig
' description   :   Company
' programmer    :   _-=[ srm ]=-_
' date          :   31 jan 2005

Option Explicit
    Dim nAdd As Integer
    Dim cSeries As String
    Dim oTempADO As New ADODB.Recordset

Private Sub Command10_Click()
    On Error GoTo ErrODBCSave
    Dim cString As String
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Company Information file entry?", vbYesNoCancel, "Supplier File Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("PA66220", "ODBCCODE='" & Text1.Text & "'") Then
                    MsgBox "Company ID already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "PA66220"), oTempADO, True
                    Script2File InsertFields(Me, "PA66220")
                    
                    Log2Audit Name, "ADD " & "ODBC ID -->" & Trim(Text1.Text)
                End If
            Else
                OpenQueryDNS EditField(Me, "PA66220", "PA66220.ODBCCODE='" & Text1.Text & "'"), oTempADO, True
                Script2File EditField(Me, "PA66220", "PA66220.ODBCCODE='" & Text1.Text & "'")
                
                Log2Audit Name, "EDIT " & "ODBC ID -->" & Trim(Text1.Text)
            End If
        Case vbCancel
            GoTo endsave
    End Select

    Lock2User Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050322
    
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "ODBCCODE='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO
    
endsave:
    Exit Sub
    
ErrODBCSave:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
End Sub

Private Sub Command11_Click()
    If nAdd = 0 Then
        Unload Me
    Else
    
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo, App.Title) = vbYes Then
            Lock2User Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050329
            If Text1.Text <> cSeries Then ResetSeries "ODBC", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            GetFields Me, oTempADO
        End If
        
    End If
    
End Sub

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then dbNavigator Screen.ActiveControl, Me, oTempADO
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    Frame2.Enabled = False
    frmLookup.showPopup 23
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "ODBCCODE='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then GetFields Me, oTempADO
    End If
    Frame2.Enabled = True
End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    cSeries = GenerateSeries("ODBC")
    While IfExists("PA66220", "PA66220.ODBCCODE='" & PadStr(cSeries, "0", Text1.MaxLength) & "'")
        cSeries = GenerateSeries("ODBC")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    
    Text1.SetFocus
End Sub

Private Sub Command8_Click()
    ' --> 20050329
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Name, Text1.ToolTipText, Text1.Text, True
        nAdd = 2
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        Text1.Enabled = False
        Text2.SetFocus
    End If
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrODBCDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, App.Title) = vbYes Then
        OpenQueryDNS "DELETE FROM PA66220 WHERE ODBCCODE='" & Text1.Text & "'", oTempADO, True
        Log2Audit Name, "DELETE " & Trim(Text1.Text) & "-" & Trim(Text2.Text)
        
        Script2File "DELETE FROM PA66220 WHERE ODBCCODE='" & Text1.Text & "'"
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
'        OpenQueryDNS "SELECT * FROM PA66220 ORDER BY ODBCCODE", oTempADO, False
        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
    End If
    
    Exit Sub
    
ErrODBCDelete:
    ErrorMsg Err.Number, Err.Description, "Delete Button", Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
    Tag = nAccess_Tag
    nAdd = 0
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    OpenQueryDNS "SELECT * FROM PA66220 ORDER BY ODBCCODE", oTempADO, False
    GetFields Me, oTempADO
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (nAdd > 0) Then
        MsgBox "Please click CANCEL to abort this entry...", vbOKOnly, App.Title
        Cancel = 1
    Else
        Log2Audit Name, "CLOSE"
    End If
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub




