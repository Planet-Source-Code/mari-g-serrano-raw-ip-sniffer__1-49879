VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "Configuración"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3195
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Preferences"
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton Option3 
         Caption         =   "Trojans"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Own"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "IANA"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox chkports 
      Caption         =   "See protocol Names:"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkports_Click()
Frame1.Visible = Not Frame1.Visible
End Sub

Private Sub Command1_Click()
frmMain.bPorts = IIf(chkports.Value = vbChecked, True, False)
If Option1.Value Then
    frmMain.Prefs = eIana
ElseIf Option2.Value Then
    frmMain.Prefs = ePredef
Else
    frmMain.Prefs = eTrojans
End If
Unload Me
End Sub
