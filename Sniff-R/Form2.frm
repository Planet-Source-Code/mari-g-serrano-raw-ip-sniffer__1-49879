VERSION 5.00
Begin VB.Form frmFiltro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtros"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   13
      Text            =   "[NAME]"
      Top             =   60
      Width           =   2715
   End
   Begin VB.OptionButton Option2 
      Caption         =   "OR"
      Height          =   195
      Left            =   1860
      TabIndex        =   12
      Top             =   1200
      Width           =   675
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "AND"
      Height          =   195
      Left            =   960
      TabIndex        =   11
      Top             =   1200
      Value           =   -1  'True
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2820
      TabIndex        =   8
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   540
      Width           =   4035
      Begin VB.CommandButton Command2 
         Caption         =   "Resolve"
         Height          =   315
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtSrcAddr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   900
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkFiltroE 
         Appearance      =   0  'Flat
         Caption         =   "Incoming Filter"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   60
         TabIndex        =   5
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Src Ip:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   4035
      Begin VB.CommandButton Command3 
         Caption         =   "Resolve"
         Height          =   315
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkFiltroS 
         Appearance      =   0  'Flat
         Caption         =   "Outgoing Filter"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   1875
      End
      Begin VB.TextBox txtDestAddr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Dest Ip:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
    'crear pestaña con datos introducidos
    If frmMain.dFiltro.Exists(CStr(txtNombre.Text)) Then
        MsgBox "Filter Name already exists!"
        Exit Sub
    End If
    
    
    Set frmMain.cF = New cFilter
    
    With frmMain.cF
    If chkFiltroE.Value = vbChecked Then
        'filtrar ip de entrada por la introducida
        .bSrc = True
        .Src = txtSrcAddr.Text
    End If
    If chkFiltroS.Value = vbChecked Then
        'filtrar ip de entrada por la introducida
        .bDest = True
        .Dest = txtDestAddr.Text
    End If
    If .bSrc And .bDest Then If Option1.Value Then .bAND = True
            
    End With
        
    frmMain.ts2.Tabs.add , txtNombre, txtNombre
    
    Set frmMain.dFiltro.Item(CStr(txtNombre.Text)) = frmMain.cF
    Unload Me
End Sub

Private Sub Command2_Click()
    Command2.Caption = "Espera..."
    txtSrcAddr.Text = HostByName(txtSrcAddr.Text)
    Command2.Caption = "Resolver"
End Sub

Private Sub Command3_Click()
    Command3.Caption = "Espera..."
    txtDestAddr.Text = HostByName(txtDestAddr.Text)
    Command3.Caption = "Resolver"
End Sub
