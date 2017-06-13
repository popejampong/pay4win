VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Menu Utility"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   53
      ImageHeight     =   60
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMenu.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   1575
      Begin VB.CommandButton Command4 
         Caption         =   "&Close"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Re&set"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   1770
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Restore"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1050
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Apply"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   315
         Width           =   1335
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   9340
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
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
End
Attribute VB_Name = "frmSysMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmSysMenu
' description   :   System Menu
' programmer    :   _-=[ srm ]=-_
' date created  :   12 january 2005

Option Explicit
    Dim oTempADO As New ADODB.Recordset
    Dim lChanged As Boolean

Function chkChange() As Boolean
    Dim lEqual As Boolean
    DoEvents
    lEqual = True
    oTempADO.Requery
    oTempADO.MoveFirst
    While lEqual And (Not oTempADO.EOF)
        lEqual = oTempADO("AVAIL") = oTempADO("TAG")
        oTempADO.MoveNext
    Wend
    chkChange = lEqual
End Function

Sub ChkNode(ByVal nIndex As Integer, lChecked As Boolean)
    Dim nCtr As Integer, cSqlStmt As String
    DoEvents
    cSqlStmt = "UPDATE TMP7668 SET TAG=" & IIf(lChecked, 1, 0) & " WHERE MNUNAME=" & cQuote & TreeView1.Nodes(nIndex).Key & cQuote
    QueryTemp cSqlStmt, objdbRs, True
    If TreeView1.Nodes(nIndex).Children > 0 Then
        For nCtr = TreeView1.Nodes(nIndex).Child.Index To TreeView1.Nodes(nIndex).Child.LastSibling.Index
            DoEvents
            TreeView1.Nodes(nCtr).Checked = lChecked
            cSqlStmt = "UPDATE TMP7668 SET TAG=" & IIf(lChecked, 1, 0) & " WHERE MNUNAME=" & cQuote & TreeView1.Nodes(nCtr).Key & cQuote
            QueryTemp cSqlStmt, objdbRs, True
            ChkNode nCtr, lChecked
        Next nCtr
    End If
End Sub

Sub CreateTemp()
    On Error GoTo ErrCreate
'    QueryTemp "CREATE TABLE TMP7668([MNUNAME] CHAR(100),[PARENT] CHAR(100),[CAPTION] CHAR(100), [AVAIL] INTEGER )", oTempADO, True
    oTempConn.Execute "CREATE TABLE TMP7668([MNUNAME] CHAR(100),[PARENT] CHAR(100),[CAPTION] CHAR(100), [AVAIL] INTEGER, [TAG] INTEGER )"
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM TMP7668", oTempADO, True
End Sub

Private Sub Command1_Click()
    Dim cSqlStmt As String
    If MsgBox("Warning!!!" & Chr$(13) & Chr$(10) & "After applying changes you cannot restore the original values." & Chr$(13) & Chr$(10) & "Save changes made?", _
        vbYesNo, App.Title) = vbYes Then
        oTempADO.MoveFirst
        
        DoEvents
        
        ShowProgress 0, , oTempADO.RecordCount
        
        While Not oTempADO.EOF
        
            'ShowProgress 2, oTempADO.AbsolutePosition, , , "Updating status of " & Trim(oTempADO("CAPTION"))
            
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100, , , "Updating status of " & Trim(oTempADO("CAPTION"))
                
            cSqlStmt = "UPDATE PA7668 SET AVAIL=" & oTempADO("TAG") & " WHERE MNUNAME=" & cQuote & oTempADO("MNUNAME") & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, True
            
            oTempADO.MoveNext
        Wend
        
        ShowProgress 4
        
        QueryTemp "UPDATE TMP7668 SET AVAIL=TAG", oTempADO, True
        Command1.Enabled = Not chkChange
    End If
End Sub

Private Sub Command2_Click()
    Dim nCtr As Integer
    
    If MsgBox("Restore original values?", vbYesNo, App.Title) = vbYes Then
        
        ShowProgress 0, , TreeView1.Nodes.Count
        DoEvents
            
        For nCtr = 2 To TreeView1.Nodes.Count
        
            ShowProgress 2, nCtr, , , "Resetting status of " & Trim(TreeView1.Nodes(nCtr).Text)
            
            oTempADO.Requery adAsyncFetch
            oTempADO.Find "MNUNAME='" & PadStr(TreeView1.Nodes(nCtr).Key, " ", 100, PadRight) & "'"
            
            TreeView1.Nodes(nCtr).Checked = oTempADO("AVAIL")
            
        Next nCtr
        
        QueryTemp "UPDATE TMP7668 SET TAG=AVAIL", oTempADO, True
        Command1.Enabled = False
        
        ShowProgress 4
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oNode As Node, cSqlStmt As String
    
    CreateTemp
    
    Set oNode = TreeView1.Nodes.Add(, , "PATMS", App.Title, 1, 1)
'    Set oNode = TreeView1.Nodes.Add("CAS", tvwChild, "mnuFileMaintenance", ConcatStr("&File Maintenance"))
'    OpenQueryDNS "SELECT DISTINCT MNUNAME,PARENT,CAPTION FROM DI7668 WHERE PARENT = 'CAS'", objdbRs, False
    OpenQueryDNS "SELECT * FROM PA7668", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        While Not objdbRs.EOF
            cSqlStmt = "INSERT INTO TMP7668([MNUNAME],[PARENT],[CAPTION],[AVAIL],[TAG])VALUES(" & _
                       "'" & objdbRs("MNUNAME") & "'," & _
                       "'" & objdbRs("PARENT") & "'," & _
                       "'" & objdbRs("CAPTION") & "'," & _
                       objdbRs("AVAIL") & "," & _
                       objdbRs("AVAIL") & ")"
            QueryTemp cSqlStmt, oTempADO, True
            Set oNode = TreeView1.Nodes.Add(Trim(objdbRs("PARENT")), tvwChild, Trim(objdbRs("MNUNAME")), Trim(ConcatStr(objdbRs("CAPTION"), "&")))
            oNode.Checked = objdbRs("AVAIL") = 1
            objdbRs.MoveNext
        Wend
    End If
    oNode.EnsureVisible
    QueryTemp "SELECT * FROM TMP7668", oTempADO, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oTempADO = Nothing
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    DoEvents
    
    TreeView1.Nodes(Node.Index).Selected = True
   
    ChkNode Node.Index, TreeView1.SelectedItem.Checked
    
    Command1.Enabled = Not chkChange
End Sub
