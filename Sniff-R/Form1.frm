VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "IP  Sniffer"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar stB 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   6570
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin Sniffer.TP_WinSplit sp 
      Height          =   6135
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   10821
      Orientation     =   1
      Ind_Color       =   -2147483635
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   60
         TabIndex        =   10
         Tag             =   "HI"
         Top             =   60
         Width           =   10935
         Begin MSComctlLib.TabStrip ts2 
            Height          =   375
            Left            =   30
            TabIndex        =   11
            Top             =   2460
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   661
            MultiRow        =   -1  'True
            HotTracking     =   -1  'True
            Placement       =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   4
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "All"
                  ImageVarType    =   2
                  ImageIndex      =   7
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Incoming"
                  ImageVarType    =   2
                  ImageIndex      =   12
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Outgoing"
                  ImageVarType    =   2
                  ImageIndex      =   4
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "IP"
                  ImageVarType    =   2
                  ImageIndex      =   8
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lv 
            Height          =   1920
            Left            =   60
            TabIndex        =   12
            Tag             =   "LO"
            Top             =   360
            Width           =   10470
            _ExtentX        =   18468
            _ExtentY        =   3387
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Protocol"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Source Addr"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Dest Addr"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Src Port"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Dest Port"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Tamaño"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Hora"
               Object.Width           =   2646
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   2722
         Left            =   150
         TabIndex        =   3
         Tag             =   "LO"
         Top             =   3263
         Width           =   10875
         Begin VB.TextBox txtData 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Text            =   "Form1.frx":08CA
            Top             =   0
            Visible         =   0   'False
            Width           =   1995
         End
         Begin MSComctlLib.TabStrip ts 
            Height          =   375
            Left            =   0
            TabIndex        =   4
            Top             =   2220
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   661
            MultiRow        =   -1  'True
            HotTracking     =   -1  'True
            Placement       =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Hex"
                  ImageVarType    =   2
                  ImageIndex      =   8
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Text"
                  ImageVarType    =   2
                  ImageIndex      =   13
               EndProperty
            EndProperty
         End
         Begin Sniffer.TP_WinSplit spBottom 
            Height          =   2235
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   3942
            Ind_Color       =   -2147483635
            Begin MSComctlLib.TreeView tv 
               Height          =   1935
               Left            =   150
               TabIndex        =   8
               Tag             =   "LO"
               Top             =   150
               Width           =   5093
               _ExtentX        =   8996
               _ExtentY        =   3413
               _Version        =   393217
               LabelEdit       =   1
               Style           =   7
               ImageList       =   "ImageList1"
               BorderStyle     =   1
               Appearance      =   0
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               Height          =   1935
               Left            =   5603
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Tag             =   "HI"
               Text            =   "Form1.frx":08D0
               Top             =   150
               Width           =   5062
            End
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":140A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":24D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C472
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1640C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":203A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A340
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":342DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":34874
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":34E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":36090
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3696A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tB 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Play"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pause"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clear"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filter"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Settings"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Save"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Open"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Menu PopUp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Filter"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sc          As cSubCls
Implements iSubCls

Private Cnt As Long
Private bPaused As Boolean
Private dicData As New Scripting.Dictionary

Private dicAddr As New Scripting.Dictionary
Public dFiltro As New Scripting.Dictionary
Public cF As cFilter

Private dIps As New Scripting.Dictionary
Private dIpsd As New Scripting.Dictionary
Private dPorts As New Scripting.Dictionary
Private dPortsD As New Scripting.Dictionary

Public bPorts As Boolean

Private dicIANA As New Scripting.Dictionary
Private dicTrojans As New Scripting.Dictionary
Private dicPredef As New Scripting.Dictionary 'los propios
Public Enum ePref
    eIana
    ePredef
    eTrojans
End Enum
Public Prefs As ePref

Private Sub LoadPortNames()
    'abrir archivos de puertos y guardar en diccionarios
    Dim ff As Integer
    Dim Linea As String
    Dim arr As Variant
    ff = FreeFile
    Open App.Path & "\iana.txt" For Input As ff
    Do While Not EOF(ff)
        Line Input #ff, Linea
        'desc   puerto  desc larga
        arr = Split(Linea, vbTab)
        dicIANA.Item(CStr(arr(1))) = arr(0) & "     " & arr(2) '...
    Loop
    Close ff
    
    ff = FreeFile
    Open App.Path & "\predefs.txt" For Input As ff
    Do While Not EOF(ff)
        Line Input #ff, Linea
        arr = Split(Linea, vbTab)
        dicPredef.Item(CStr(arr(0))) = arr(1)
    Loop
    Close ff
    
        ff = FreeFile
    Open App.Path & "\trojans.txt" For Input As ff
 '   On Error GoTo e
    Do While Not EOF(ff)
        Line Input #ff, Linea
        arr = Split(Linea, vbTab)
        dicTrojans.Item(CStr(arr(0))) = arr(1)
    Loop
    Close ff
'e:
   ' Resume
End Sub

Private Function GetPortName(p1 As String) As String
Select Case Prefs
Case eIana
    If dicIANA.Exists(p1) Then
        GetPortName = dicIANA.Item(p1)
    Else
       GetPortName = p1
    End If
Case ePredef
    If dicPredef.Exists(p1) Then
        GetPortName = dicPredef.Item(p1)
    Else
       GetPortName = p1
    End If
Case eTrojans
    If dicTrojans.Exists(p1) Then
        GetPortName = dicTrojans.Item(p1)
    Else
       GetPortName = p1
    End If
End Select
End Function


Private Sub Form_Load()
    On Error GoTo errH
    LoadPortNames
    DisplayAdapterInfo

    tv.Nodes.add , , "S", "Source", 9
    tv.Nodes.add , , "D", "Dest", 13


    Set sc = New cSubCls
    sc.AddMsg WINSOCKMSG
    sc.SubClass Me.hWnd, Me
Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    sp.Move 0, tB.Height, ScaleWidth, ScaleHeight - stB.Height - tB.Height
    sp_Sized
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errhand
    Call WSAAsyncSelect(lSocket, Me.hWnd, ByVal 1025, 0)
    closesocket lSocket
    EndWinsock
    sc.UnSubclass
Exit Sub
errhand:
    MsgBox Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError, vbCritical, "wsck_displayadapterinfo"
End Sub

Private Sub DisplayAdapterInfo()
On Error GoTo errhand
Combo1.Clear
Dim str() As String
Call wsck_enum_interfaces(str)
Dim i As Integer
Dim v As Variant
For i = 0 To UBound(str)
    v = Split(str(i), ";")
    If v(0) <> "127.0.0.1" Then
        Combo1.AddItem v(0)
        dicAddr.Item(v(0)) = v(0)
    End If
Next i
Combo1.ListIndex = 0
Exit Sub
errhand:
    MsgBox Err.Description
End Sub


Private Sub iSubCls_Antes(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'
End Sub

'Procesamos el único mensaje que llega
Private Sub iSubCls_Despues(lReturn As Long, ByVal hWnd As Long, _
        ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    
    Const Size As Long = 2000
    Dim X      As Long
    Dim tLen   As Long
    Dim OffSet As Long
    Dim ReadBuffer(0 To Size - 1) As Byte
    Dim ip_header  As ipheader
    Dim tcp_header As tcpheader
    Dim udp_header As udpheader
    Dim sBytes   As String, _
        SrcAddr  As String, _
        DestAddr As String, _
        SrcPort  As String, _
        DestPort As String, _
        Proto    As String, _
        strBuff  As String
    
    If lParam = FD_READ Then 'hay datos esperando a ser leidos
       Do
         X = recv(wParam, ReadBuffer(0), Size, 0)
         If X Then
            Cnt = Cnt + 1
            stB.Panels(1).Text = Cnt
            CopyMemory ip_header, ReadBuffer(0), Len(ip_header)
               
            sBytes = ntohs(ip_header.ip_totallength) & " bytes. "
            SrcAddr = GetAscIp(ip_header.ip_srcaddr)
            DestAddr = GetAscIp(ip_header.ip_destaddr)
            'ICMP
            If ip_header.ip_protocol = 1 Then
               Proto = "ICMP"
               OffSet = 20 'IP?
            'TCP
            ElseIf ip_header.ip_protocol = 6 Then
               Proto = "TCP"                        'Len(ip_header)
               CopyMemory tcp_header, ReadBuffer(0 + 20), 20 ' Len(tcp_header)
              
               SrcPort = ntohs(tcp_header.src_portno)
               DestPort = ntohs(tcp_header.dst_portno)
               OffSet = 40 'IP+TCP
               
            'UDP
            ElseIf ip_header.ip_protocol = 17 Then
               Proto = "UDP"                         'Len(ip_header)
               CopyMemory udp_header, ReadBuffer(0 + 20), 8 'Len(udp_header)
              
               SrcPort = ntohs(udp_header.src_portno)
               DestPort = ntohs(udp_header.dst_portno)
                OffSet = 28 'IP+UDP
            Else
                Proto = ip_header.ip_protocol
                OffSet = 20 'IP?
            End If
            
            tLen = X - OffSet
            'x es lo q lee.
            If tLen Then
               strBuff = Space$(tLen)                 'nos saltamos los headers.
               CopyMemory ByVal strBuff, ByVal VarPtr(ReadBuffer(OffSet)), tLen
            End If
            Add2Lv Proto, SrcAddr, DestAddr, SrcPort, DestPort, strBuff, tLen
         End If 'X
       Loop Until X <> Size
   Else
   Debug.Print uMsg
    
    End If
End Sub

Private Sub Add2Lv(Proto As String, SrcAddr As String, DestAddr As String, _
        SrcPort As String, DestPort As String, Data As String, Tam As Long)
'si ya existe origen, añadimos
'si origen es combo1: flecha roja
'si destino=combo1:flecha verde
Dim clV As clV
Dim Hora As String

Set clV = New clV

Hora = Time$ & Right$(Format$(Timer, "#.00"), 3)
clV.add Proto, SrcAddr, SrcPort, DestAddr, DestPort, CStr(Tam), Data, Hora

Static i As Long
i = i + 1
Set dicData.Item(CStr(i)) = clV

Select Case ts2.SelectedItem.Index

Case 1 'todo

Case 2 'solo recibido
    If dicAddr.Exists(SrcAddr) Then Exit Sub

Case 3 'solo enviado
    If dicAddr.Exists(DestAddr) Then Exit Sub
Case 4 'chequear si ya se ha metido un Src-Dest igual
    Dim j As Long
    For j = 1 To lv.ListItems.Count - 1 'el source ya está?
       If lv.ListItems(j).ListSubItems(1).Text = dicData.Item(CStr(i)).mSrcAddr Then
         If lv.ListItems(j).ListSubItems(2).Text = dicData.Item(CStr(i)).mDestAddr Then
           Exit Sub
         End If
       End If
    Next
Case Else
    ' depende del filtro
    Dim bSrc As Boolean, bDest As Boolean, bDo As Boolean
    Set cF = dFiltro.Item(ts2.SelectedItem.Caption)
    With cF
        If .bSrc Then
            If SrcAddr = .Src Then bSrc = True
        End If
        If .bDest Then
            If DestAddr = .Dest Then bDest = True
        End If
        If .bAND Then
            bDo = bSrc And bDest
        Else
            bDo = bSrc Or bDest
        End If
        If Not bDo Then Exit Sub
    End With
    
End Select
    
AddLV clV
End Sub

Private Sub addTv(SrcAddr$, SrcPort$, DestAddr$, DestPort$)


If dIps.Exists(SrcAddr) Then 'existe esta ip de src
   If dIps.Item(SrcAddr).Exists(CStr(SrcPort)) Then 'existe este puerto de origen
    'ok.
   Else
    'se mete puerto
     dPorts.Item(CStr(SrcPort)) = "S"
     Set dIps.Item(SrcAddr) = dPorts '.Item(CStr(SrcPort))
     
    tv.Nodes.add "S" & SrcAddr, tvwChild, , SrcPort
   
   End If
   
Else
   'se mete ip y puerto
   dPorts.Item(CStr(SrcPort)) = "S"
   Set dIps.Item(SrcAddr) = dPorts '.Item(CStr(SrcPort))
   
   tv.Nodes.add "S", tvwChild, "S" & SrcAddr, SrcAddr
   tv.Nodes.add "S" & SrcAddr, tvwChild, , SrcPort
End If

If dIpsd.Exists(DestAddr) Then 'existe esta ip de src
   If dIpsd.Item(DestAddr).Exists(CStr(DestPort)) Then 'existe este puerto de origen
    'ok.
   Else
    'se mete puerto
     dPortsD.Item(CStr(DestPort)) = "S"
     Set dIpsd.Item(DestAddr) = dPortsD '.Item(CStr(DestPort))
     
     tv.Nodes.add "D" & DestAddr, tvwChild, , DestPort
   End If
   
Else
   'se mete ip y puerto
   dPortsD.Item(CStr(DestPort)) = "S"
   Set dIpsd.Item(DestAddr) = dPortsD '.Item(CStr(DestPort))
   
   tv.Nodes.add "D", tvwChild, "D" & DestAddr, DestAddr
   tv.Nodes.add "D" & DestAddr, tvwChild, , DestPort
End If



End Sub

Private Sub AddLV(c As clV)
Dim li As ListItem
    addTv c.mSrcAddr, c.mSrcPort, c.mDestAddr, c.mDestPort
    

    If dicAddr.Exists(c.mSrcAddr) Then    'enviamos
        Set li = lv.ListItems.add(, , c.mProtocol, , 4)
    Else 'recibimos
        Set li = lv.ListItems.add(, , c.mProtocol, , 12)
    End If
    li.Tag = c.mData
    
    With li.ListSubItems
        .add , , c.mSrcAddr
        .add , , c.mDestAddr
    If Me.bPorts Then
        .add(, , GetPortName(c.mSrcPort)).Tag = c.mSrcPort
        .add(, , GetPortName(c.mDestPort)).Tag = c.mDestPort
    Else
        .add(, , c.mSrcPort).Tag = c.mSrcPort
        .add(, , c.mDestPort).Tag = c.mDestPort
    End If
    .add , , c.mSize
    .add , , c.mHora
    End With


If Not bPaused Then
    li.EnsureVisible
    li.Selected = True
End If
End Sub



Private Sub lv_DblClick()
    'mostrar informacion de puertos:nombres
    frmInfo.Label1.Caption = "   " & _
    lv.SelectedItem.ListSubItems(3).Tag & vbCrLf & " [IANA]: " & _
    dicIANA.Item(lv.SelectedItem.ListSubItems(3).Tag) & vbCrLf & _
    " [PREDEF]: " & _
    dicPredef.Item(lv.SelectedItem.ListSubItems(3).Tag) & vbCrLf & _
    " [TROJAN]: " & _
    dicTrojans.Item(lv.SelectedItem.ListSubItems(3).Tag) & vbCrLf & _
    "   " & lv.SelectedItem.ListSubItems(4).Tag & vbCrLf & " [IANA]: " & _
    dicIANA.Item(lv.SelectedItem.ListSubItems(4).Tag) & vbCrLf & _
    " [PREDEF]: " & _
    dicPredef.Item(lv.SelectedItem.ListSubItems(4).Tag) & vbCrLf & _
    " [TROJAN]: " & _
    dicTrojans.Item(lv.SelectedItem.ListSubItems(4).Tag)

    frmInfo.Show
End Sub

Private Sub lV_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim c As Long, i As Long
Dim s As String, sH As String, sT As String
Dim v As String





    Dim Part As Long
    Part = (((Text1.Width / 105) - 6) / 4) - 1
    Text1.Text = vbNullString
    txtData.Text = vbNullString
    For i = 1 To Len(Item.Tag)
        c = c + 1
        v = Mid$(Item.Tag, i, 1)
        s = Hex(Asc(v))
        If Len(s) = 1 Then s = "0" & s
        sH = sH & " " & s
        s = v
        If Asc(s) < 32 Then s = "."
        sT = sT & s

        If c = Part Then
            c = 0
            txtData.Text = txtData.Text & sT
            Text1.Text = Text1.Text & " " & sH & " | " & sT & vbCrLf
            sH = vbNullString
            sT = vbNullString
        End If

    Next i
    'If ts.SelectedItem.Index = 1 Then
        
    'Else
        If LenB(sH) > 0 Then
            txtData.Text = txtData.Text & sT
            Text1.Text = Text1.Text & " " & sH & Space$((Part * 3) - Len(sH)) & " | " & sT & vbCrLf
        End If
    'End If
End Sub

Private Sub mnuDelete_Click()
    ts2.Tabs.Remove (ts2.SelectedItem.Index)
End Sub

Private Sub sp_Sized()
    spBottom.Move 0, 0, Frame1.Width, Frame1.Height - ts.Height
    txtData.Move 0, 0, Frame1.Width, Frame1.Height - ts.Height
    ts.Width = Frame1.Width
    ts.Top = spBottom.Height
    
    lv.Move 0, 0, Frame2.Width, Frame2.Height - ts2.Height
    ts2.Width = Frame2.Width
    ts2.Top = lv.Height
    
End Sub

Private Sub tB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 'Play
    Call sPlay
Case 2 'Pause
   Call sPause
Case 3 'Stop
   Call sStop
Case 4 'Delete
    lv.ListItems.Clear
Case 5 'filtro
    frmFiltro.Show
Case 6 'config
    frmConfig.Show
Case 7 'save
Case 8 'open
End Select

End Sub
Private Sub sPlay()
    On Error GoTo errH
    Cnt = 0
    If Not IsWindowsNT5 Then
        MsgBox "¡NT5 or Higher!"
        Exit Sub
    End If
    StartWinsock
    lSocket = ConnectSock(Combo1.Text, 7000, Me.hWnd)
    tB.Buttons(1).Enabled = False
    tB.Buttons(2).Enabled = True
    tB.Buttons(3).Enabled = True
    Combo1.Enabled = False
    Exit Sub
errH:
    MsgBox Err.Description
End Sub
Private Sub sPause()
    bPaused = Not bPaused

End Sub

Private Sub sStop()
    On Error GoTo errH
    Call WSAAsyncSelect(lSocket, Me.hWnd, ByVal 1025, 0)
    closesocket lSocket
    EndWinsock
    tB.Buttons(1).Enabled = True
    tB.Buttons(2).Enabled = False
    tB.Buttons(3).Enabled = False
    Combo1.Enabled = True
    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub ts_Click()
Select Case ts.SelectedItem.Index
Case 1
    'Paquetes
    spBottom.Visible = True
    txtData.Visible = False
Case 2
    'Datos
    spBottom.Visible = False
    txtData.Visible = True
End Select
End Sub

Private Sub ts2_Click()
    Dim i As Long, j As Long
    lv.ListItems.Clear
    tv.Nodes.Clear
    dIps.RemoveAll
    dPorts.RemoveAll
    dIpsd.RemoveAll
    dPortsD.RemoveAll
    tv.Nodes.add , , "S", "Source", 9
    tv.Nodes.add , , "D", "Dest", 13
Select Case ts2.SelectedItem.Index
Case 1 'todo
    For i = 1 To dicData.Count - 1
        AddLV dicData.Item(CStr(i))
    Next
Case 2 'recibido
    For i = 1 To dicData.Count - 1
        If dicAddr.Exists(dicData.Item(CStr(i)).mDestAddr) Then AddLV dicData.Item(CStr(i))
    Next

Case 3 'enviado
    For i = 1 To dicData.Count - 1
        If dicAddr.Exists(dicData.Item(CStr(i)).mSrcAddr) Then AddLV dicData.Item(CStr(i))
    Next
Case 4
    'añadir solo 1 ip
    Dim bYa As Boolean
    For i = 1 To dicData.Count - 1
        For j = 1 To lv.ListItems.Count - 1 'el source ya está?
            If lv.ListItems(j).ListSubItems(1).Text = dicData.Item(CStr(i)).mSrcAddr Then
                'ya esta
              
                If lv.ListItems(j).ListSubItems(2).Text = dicData.Item(CStr(i)).mDestAddr Then
                    bYa = True
                    Exit For
                End If
            End If
        Next
        If Not bYa Then AddLV dicData.Item(CStr(i))
        
        
        
    Next
    
    
Case Else
    'segun el filtro de la pestaña...
    Dim bSrc As Boolean, bDest As Boolean, bDo As Boolean
    For i = 1 To dicData.Count - 1
        bSrc = False
        bDest = False
        bDo = False
        Set cF = dFiltro.Item(ts2.SelectedItem.Caption)
        
        With cF
            If .bSrc Then
                If dicData.Item(CStr(i)).mSrcAddr = .Src Then bSrc = True
            End If
            If .bDest Then
                If dicData.Item(CStr(i)).mDestAddr = .Dest Then bDest = True
            End If
            If .bAND Then
               bDo = bSrc And bDest
            Else
                bDo = bSrc Or bDest
            End If
            
            If bDo Then AddLV dicData.Item(CStr(i))
        End With

    Next


End Select
End Sub

Private Sub ts2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ts2.SelectedItem.Index > 4 Then
    If Button = vbRightButton Then
        PopupMenu PopUp
    End If
End If

End Sub
