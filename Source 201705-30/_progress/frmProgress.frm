VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3540
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmProgress.frx":0000
   ScaleHeight     =   3540
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   3165
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.ListBox List1 
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
      Height          =   1785
      Left            =   90
      TabIndex        =   6
      Top             =   1020
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   -30
      TabIndex        =   5
      Top             =   870
      Width           =   4860
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   60
      TabIndex        =   4
      Top             =   2895
      Width           =   4635
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait processing..."
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   2955
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Processing 1 of 1000 records..."
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   75
      TabIndex        =   1
      Top             =   315
      Width           =   3900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dong-In Cost Accounting System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   4080
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
