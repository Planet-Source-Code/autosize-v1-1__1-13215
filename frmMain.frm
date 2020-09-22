VERSION 5.00
Object = "{37DF34C6-C510-11D4-AF97-0060973144DB}#1.0#0"; "autosize.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   1710
   ClientTop       =   1995
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   6495
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   3300
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   60
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   60
      Width           =   3135
   End
   Begin autosize.ControlSizer ControlSizer1 
      Left            =   660
      Top             =   2940
      _ExtentX        =   953
      _ExtentY        =   873
   End
   Begin autosize.FormSizer FormSizer1 
      Left            =   120
      Top             =   2940
      _ExtentX        =   926
      _ExtentY        =   873
      ProgramName     =   "Defualt"
   End
   Begin VB.TextBox Text3 
      Height          =   3075
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMain.frx":0000
      Top             =   420
      Width           =   6375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  ControlSizer1.ClearRelations
  Call ControlSizer1.AddRelation(Me, csrHCenter, Text1, csrWidth, -120, True)
  Call ControlSizer1.AddRelation(Text1, csrRight, Text2, csrLeft, 120, False)
  Call ControlSizer1.AddRelation(Me, csrRight, Text2, csrWidth, -180, True)
  Call ControlSizer1.AddRelation(Me, csrRight, Text3, csrWidth, -180, True)
  Call ControlSizer1.AddRelation(Me, csrBottom, Text3, csrHeight, -465, True)
End Sub

