VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAccess 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Access Rights"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   720
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   10
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   3840
      TabIndex        =   4
      Top             =   2160
      Width           =   3135
      Begin VB.CommandButton Command4 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   615
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Restore"
         Height          =   615
         Left            =   1080
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Apply"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   375
      Left            =   1620
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4230
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7461
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   3840
      TabIndex        =   7
      Top             =   960
      Width           =   3135
      Begin VB.CheckBox Check1 
         Caption         =   "Delete"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Edit"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Access Rights"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   675
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   3840
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   675
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   240
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "frmAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmDeduction
' description   :   Module for Maintenance of Deduction
' programmer    :   _-[srm]-_
' date          :   17 Oct 2005
' note          :   copied from DICAS...


Option Explicit
    Dim oTempADO As New ADODB.Recordset
    Dim lChanged As Boolean

Function IsChange() As Boolean
    Dim lEqual As Boolean, nCtr As Integer
    lEqual = True
    
    While lEqual And nCtr < MSFlexGrid1.Rows - 1
        DoEvents
        nCtr = nCtr + 1
        lEqual = (MSFlexGrid1.TextMatrix(nCtr, 2) = MSFlexGrid1.TextMatrix(nCtr, 5)) And _
                 (MSFlexGrid1.TextMatrix(nCtr, 3) = MSFlexGrid1.TextMatrix(nCtr, 6)) And _
                 (MSFlexGrid1.TextMatrix(nCtr, 4) = MSFlexGrid1.TextMatrix(nCtr, 7)) And _
                 (MSFlexGrid1.TextMatrix(nCtr, 8) = MSFlexGrid1.TextMatrix(nCtr, 9))
    Wend
    IsChange = lEqual
End Function

Sub ClearSetting()
    Dim nX As Integer, nY As Integer
    For nX = 1 To MSFlexGrid1.Rows - 1
        For nY = 2 To 9
            MSFlexGrid1.TextMatrix(nX, nY) = 0
        Next nY
    Next nX
    TreeView1.Nodes(1).Checked = False
    ChkNode 1, False
End Sub

Sub LoadSetting()
    Dim nPos As Integer
        
    ShowProgress 0, , , "Please wait..."
        
    OpenQueryDNS "SELECT * FROM PA2798 WHERE USERID=" & cQuote & Text1.Text & cQuote, oTempADO, False
    If oTempADO.RecordCount > 0 Then
    
        ShowProgress 1, , oTempADO.RecordCount
        
        While Not oTempADO.EOF
        
            'ShowProgress 2, frmProgress.ProgressBar1.Value + 1, , "loading access right for " & oTempADO("MNUNAME")
            
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100, , "loading access right for " & oTempADO("MNUNAME")
            
            nPos = ChkGrid(oTempADO("MNUNAME"))
            If nPos > 0 Then
                MSFlexGrid1.TextMatrix(nPos, 2) = oTempADO("BIT1")
                MSFlexGrid1.TextMatrix(nPos, 3) = oTempADO("BIT2")
                MSFlexGrid1.TextMatrix(nPos, 4) = oTempADO("BIT3")
                MSFlexGrid1.TextMatrix(nPos, 8) = 1
                MSFlexGrid1.TextMatrix(nPos, 5) = oTempADO("BIT1")
                MSFlexGrid1.TextMatrix(nPos, 6) = oTempADO("BIT2")
                MSFlexGrid1.TextMatrix(nPos, 7) = oTempADO("BIT3")
                MSFlexGrid1.TextMatrix(nPos, 9) = 1
            End If
            
            For nPos = 1 To TreeView1.Nodes.Count
                If TreeView1.Nodes(nPos).Key = oTempADO("MNUNAME") Then
                    TreeView1.Nodes(nPos).Checked = True
                    Exit For
                End If
            Next nPos
            
            oTempADO.MoveNext
        Wend
        
    End If
    
    ShowProgress 4
End Sub

Function ChkGrid(ByVal cKey As String) As Integer
    Dim nCtr As Integer
    ChkGrid = 0
    DoEvents
    For nCtr = 1 To MSFlexGrid1.Rows - 1
        If cKey = MSFlexGrid1.TextMatrix(nCtr, 1) Then
            ChkGrid = nCtr
            Exit For
        End If
    Next nCtr
End Function

Sub ChkNode(ByVal nIndex As Integer, lChecked As Boolean, Optional ByVal nCheck As Integer = 3)
    Dim nCtr As Integer, nPos As Integer, cSqlStmt As String
    DoEvents
    
    nPos = ChkGrid(TreeView1.Nodes(nIndex).Key)
    If nPos > 0 Then
        If nCheck = 3 Then
            MSFlexGrid1.TextMatrix(nPos, 9) = IIf(lChecked, 1, 0)
        Else
            MSFlexGrid1.TextMatrix(nPos, 5 + nCheck) = IIf(lChecked, 1, 0)
        End If
    End If
    
    If TreeView1.Nodes(nIndex).Children > 0 Then
        For nCtr = TreeView1.Nodes(nIndex).Child.Index To TreeView1.Nodes(nIndex).Child.LastSibling.Index
            DoEvents
            If nCheck = 3 Then
                TreeView1.Nodes(nCtr).Checked = lChecked
            End If
            ChkNode nCtr, lChecked, nCheck
        Next nCtr
    End If
End Sub

Private Sub Check1_Click(Index As Integer)
    On Error GoTo ErrChk
    If TypeOf ActiveControl Is CheckBox Then
        ChkNode TreeView1.SelectedItem.Index, Check1(Index).Value, Index
        Command2.Enabled = Not IsChange
    End If
    
    Exit Sub
ErrChk:
    MsgBox "Please select a node first!", vbCritical, App.Title
End Sub

Private Sub Command1_Click()
    cResult = ""
    frmLookup.showPopup 1, " WHERE SYSUSER=1"
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "SELECT * FROM PA2360 WHERE USERID=" & cQuote & cResult & cQuote, objdbRs, False
        Text1.Text = objdbRs("USERID")
        Text2.Text = Trim(objdbRs("FIRSTNAME")) & " " & Trim(objdbRs("LASTNAME"))
        ClearSetting
        LoadSetting
    End If
End Sub

Private Sub Command2_Click()
    Dim nCtr As Integer, cSqlStmt As String
    
    If MsgBox("Warning!!!" & Chr$(13) & Chr$(10) & "After applying changes you cannot restore the original values." & Chr$(13) & Chr$(10) & "Save changes made?", _
        vbYesNo, App.Title) = vbYes Then
        
        ShowProgress 0, , , "Please wait..."
        
        DoEvents
        
        OpenQueryDNS "DELETE FROM PA2798 WHERE USERID=" & cQuote & Text1.Text & cQuote, objdbRs, True
        
        Log2Audit Name, "Delete all menu assigned to UserID#" & Text1.Text
        
        Script2File "DELETE FROM PA2798 WHERE USERID=" & cQuote & Text1.Text & cQuote
        
        ShowProgress 1, , MSFlexGrid1.Rows - 1
        
        For nCtr = 1 To MSFlexGrid1.Rows - 1
        
            'ShowProgress 2, nCtr, , , "Updating access for " & Trim(MSFlexGrid1.TextMatrix(nCtr, 1))
            ShowProgress 2, (nCtr / (MSFlexGrid1.Rows - 1)) * 100, , , "Updating access for " & Trim(MSFlexGrid1.TextMatrix(nCtr, 1))
            
            If MSFlexGrid1.TextMatrix(nCtr, 9) = 1 Then
                cSqlStmt = "INSERT INTO PA2798(USERID,MNUNAME,BIT1,BIT2,BIT3)VALUES(" & _
                           cQuote & Text1.Text & cQuote & "," & _
                           cQuote & MSFlexGrid1.TextMatrix(nCtr, 1) & cQuote & "," & _
                           MSFlexGrid1.TextMatrix(nCtr, 5) & "," & _
                           MSFlexGrid1.TextMatrix(nCtr, 6) & "," & _
                           MSFlexGrid1.TextMatrix(nCtr, 7) & ")"
                OpenQueryDNS cSqlStmt, oTempADO, True
                
                Log2Audit Name, "Assigned " & MSFlexGrid1.TextMatrix(nCtr, 1) & " to UserID#" & Text1.Text
                Script2File cSqlStmt
        
            End If
        Next nCtr
        
        ShowProgress 4
        
        ClearSetting
        LoadSetting
        Command2.Enabled = Not IsChange
    End If
End Sub

Private Sub Command3_Click()
    If MsgBox("Restore to original settings?", vbYesNo, App.Title) = vbYes Then
        ClearSetting
        LoadSetting
        Command2.Enabled = False
    End If
End Sub

Private Sub Command4_Click()
    If Command2.Enabled Then
        If MsgBox("Changes have been made..." & Chr$(13) & Chr(10) & "Save settings?", vbYesNo, App.Title) = vbYes Then
            Command2_Click
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oNode As Node, cSqlStmt As String, nCtr As Integer
    
    Log2Audit Name, "OPEN"
    
    OpenQueryDNS "SELECT * FROM PA2360", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        Text1.Text = objdbRs("USERID")
        Text2.Text = Trim(objdbRs("FIRSTNAME")) & " " & Trim(objdbRs("LASTNAME"))
    End If
    
    Set oNode = TreeView1.Nodes.Add(, , "PATMS", App.Title)
    
    OpenQueryDNS "SELECT * FROM PA7668 WHERE AVAIL=1", objdbRs, False
    If objdbRs.RecordCount > 0 Then
    
        nCtr = 0
        MSFlexGrid1.Rows = objdbRs.RecordCount + 1
        
        While Not objdbRs.EOF
            nCtr = nCtr + 1
            
            With MSFlexGrid1
                .TextMatrix(nCtr, 1) = objdbRs("MNUNAME")
                ' --> next 3 is for add, edit, delete
                .TextMatrix(nCtr, 2) = 0
                .TextMatrix(nCtr, 3) = 0
                .TextMatrix(nCtr, 4) = 0
                ' --> next 3 is for temp add, edit, delete
                .TextMatrix(nCtr, 5) = 0
                .TextMatrix(nCtr, 6) = 0
                .TextMatrix(nCtr, 7) = 0
                ' --> last 2 is for status and availability of menu
                .TextMatrix(nCtr, 8) = 0
                .TextMatrix(nCtr, 9) = 0
            End With
            
            Set oNode = TreeView1.Nodes.Add(Trim(objdbRs("PARENT")), tvwChild, Trim(objdbRs("MNUNAME")), Trim(ConcatStr(objdbRs("CAPTION"), "&")))
            'oNode.Checked = objdbRs("AVAIL") = 1
            
            objdbRs.MoveNext
        Wend
        
    End If
    
    oNode.EnsureVisible
    
    ClearSetting
    LoadSetting
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Log2Audit Name, "CLOSE"
    Unload Me
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    If TypeOf ActiveControl Is TreeView Then
        DoEvents
        ChkNode Node.Index, TreeView1.Nodes(Node.Index).Checked
        Command2.Enabled = Not IsChange
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim nPos As Integer
    nPos = ChkGrid(Node.Key)
    If nPos > 0 Then
        Check1(0).Value = Int(MSFlexGrid1.TextMatrix(nPos, 5))
        Check1(1).Value = Int(MSFlexGrid1.TextMatrix(nPos, 6))
        Check1(2).Value = Int(MSFlexGrid1.TextMatrix(nPos, 7))
    Else
        Check1(0).Value = False
        Check1(1).Value = False
        Check1(2).Value = False
    End If
    Command2.Enabled = Not IsChange
End Sub
