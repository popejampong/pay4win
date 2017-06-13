VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSplashConfig 
   BorderStyle     =   0  'None
   ClientHeight    =   5250
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplashConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplashConfig.frx":57E2
   ScaleHeight     =   5250
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1140
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   3334
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   4210752
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   6245
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Server Configuration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "v01.000.0016"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   6285
      TabIndex        =   0
      Top             =   5025
      Width           =   1155
   End
End
Attribute VB_Name = "frmSplashConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmSplash
' description   :   Splash Screen
' programmer    :   _-=[ srm ]=-_
' date          :   17 Oct 2005
' note          :   copied from DICAS

Option Explicit

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()

    objdbConn.Close
    Set oTempConn = Nothing
    Set objdbConn = Nothing
    Set objdbRs = Nothing


    Load frmSplash
    frmSplash.Show 0
    Add2List "Loading Cost Accounting System, please wait..."
    DoEvents

    Add2List "Please wait reading from initial configuration"

    If Not ReadINI Then
        Add2List "Initial configuration not found!"
        End
    End If
    Add2List "Initial configuration successfully initiated..."

    Write2File "Session start at " & Now

    Add2List "Please wait establishing connection..."
    If Not DetectServer Then
        Add2List "Connection failed!"
        End
    End If
    Add2List "Connection successful!"

'    ChkTable

    Log2Audit "frmMain", "Open " & App.Title & " version " & PadStr(App.Major, "0", 2) & "." & PadStr(App.Minor, "0", 3) & "." & PadStr(App.Revision, "0", 4)

    Add2List "Please wait initializing temporary database"
    If Not DetectTemp Then
        Add2List "Temporary database initialization failed!"
        End
    End If

    Add2List "Accessing the system..."

    Log2Audit "frmMain", "Session started."

    Unload Me

    Load frmMain
    frmMain.showLogin

'    Load frmWebReport
'    frmWebReport.WebBrowser1.Navigate "http://localhost"
'    frmWebReport.Show

    With frmMain
        .Show
    End With
    
    
    
End Sub

Private Sub cmdOK_Click()

    Dim nCtr As Integer, _
        cParam As String
        
        cParam = ""
        For nCtr = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(nCtr).Selected Then cParam = ListView1.ListItems(nCtr).Text
        Next nCtr
            


        cODBC = ""
        cUser = ""
        cPwd = ""
        cConnString = ""


        OpenQueryDNS " select ODBCSERVER, ODBCUSER, ODBCPASSWORD, ODBCDATABASE, CMPID " & _
                     "  from pa66220 where ODBCCODE = " & cQuote & cParam & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            gServerIP = objdbRs("ODBCSERVER")
            cODBC = objdbRs("ODBCDATABASE")
            cUser = objdbRs("ODBCUSER")
            cPwd = objdbRs("ODBCPASSWORD")
        End If
        
        cConnString = "DRIVER={MySQL ODBC 3.51 Driver};" & _
                      "SERVER=" & gServerIP & ";" & _
                      "DATABASE=" & cODBC & ";" & _
                      "USER=" & cUser & ";" & _
                      "PASSWORD=" & cPwd & ";" & _
                      "OPTION=11;"
        
        
    objdbConn.Close
    Set oTempConn = Nothing
    Set objdbConn = Nothing
    Set objdbRs = Nothing


    Load frmSplash
    frmSplash.Show 0
    Add2List "Loading Cost Accounting System, please wait..."
    DoEvents

    Add2List "Please wait reading from initial configuration"

'    If Not ReadINI Then
'        Add2List "Initial configuration not found!"
'        End
'    End If
'    Add2List "Initial configuration successfully initiated..."

    Write2File "Session start at " & Now

    Add2List "Please wait establishing connection..."
    If Not DetectServer Then
        Add2List "Connection failed!"
        End
    End If
    Add2List "Connection successful!"

'    ChkTable

    Log2Audit "frmMain", "Open " & App.Title & " version " & PadStr(App.Major, "0", 2) & "." & PadStr(App.Minor, "0", 3) & "." & PadStr(App.Revision, "0", 4)

    Add2List "Please wait initializing temporary database"
    If Not DetectTemp Then
        Add2List "Temporary database initialization failed!"
        End
    End If

    Add2List "Accessing the system..."

    Log2Audit "frmMain", "Session started."

    Unload Me

    Load frmMain
    frmMain.showLogin

'    Load frmWebReport
'    frmWebReport.WebBrowser1.Navigate "http://localhost"
'    frmWebReport.Show

    With frmMain
        .Show
    End With
End Sub

Private Sub Form_Load()

    OpenQueryDNS "select ODBCCODE as ID, DESCRIPTION from pa66220 ", objdbRs, False
    add2LstBox2 objdbRs, ListView1, Array("DESCRIPTION", "ID")

    lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub



Sub add2LstBox2(ByVal oRecordSet As ADODB.Recordset, ByVal oListBox As ListView, ByVal aField As Variant)
    Dim lstItem As ListItem
    
    If oRecordSet.RecordCount > 0 Then
        oListBox.ListItems.Clear
        While Not oRecordSet.EOF
            Set lstItem = oListBox.ListItems.Add()
            lstItem.Text = objdbRs(aField(1))
            lstItem.SubItems(1) = objdbRs(aField(0))
            oRecordSet.MoveNext
        Wend
    End If
End Sub


