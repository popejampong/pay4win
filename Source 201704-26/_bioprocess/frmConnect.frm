VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Begin VB.Form frmConnect 
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   495
      Left            =   915
      TabIndex        =   1
      Top             =   3555
      Width           =   1215
   End
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   2880
      Left            =   225
      OleObjectBlob   =   "frmConnect.frx":0000
      TabIndex        =   0
      Top             =   75
      Width           =   2880
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    CZKEM1.Connect_Net "192.168.1.201", 0
    MsgBox CZKEM1.GetSerialNumber(1, "")
    
End Sub

Private Sub CZKEM1_OnConnected()
    MsgBox "Test"
End Sub
