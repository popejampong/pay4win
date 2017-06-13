VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   5250
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":57E2
   ScaleHeight     =   5250
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   60
      TabIndex        =   0
      Top             =   1425
      Width           =   5295
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
      TabIndex        =   1
      Top             =   5025
      Width           =   1155
   End
End
Attribute VB_Name = "frmSplash"
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

Private Sub Form_Load()
    lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub
