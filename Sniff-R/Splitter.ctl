VERSION 5.00
Begin VB.UserControl TP_WinSplit 
   Appearance      =   0  'Flat
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   ControlContainer=   -1  'True
   DrawStyle       =   4  'Dash-Dot-Dot
   DrawWidth       =   3
   FillColor       =   &H00E2E7E9&
   ForeColor       =   &H80000006&
   ScaleHeight     =   2040
   ScaleWidth      =   1950
   ToolboxBitmap   =   "Splitter.ctx":0000
   Begin VB.PictureBox picIndicator 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   270
      ScaleHeight     =   613.975
      ScaleMode       =   0  'User
      ScaleWidth      =   14352
      TabIndex        =   0
      Tag             =   "LOCAL"
      Top             =   315
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Image imgSense 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1800
      Left            =   120
      MousePointer    =   9  'Size W E
      Tag             =   "LOCAL"
      Top             =   105
      Width           =   1695
   End
End
Attribute VB_Name = "TP_WinSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' poner HI y LO en el tag de cada Control
' control splitter (ajusta el tamaño de 2 controles arrastrando)
' MaRio Glez.Serrano.

Public Event NewSplit(ByVal NewPerCent As Long)
Public Event Sized()

''Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private sliderMoving As Boolean
Private SPLIND_CORRECT As Single
Private LastPercent As Single
Private xRefOffst As Single
Private yRefOffst As Single

'  variable para cambiar el tamaño en el IDE
'  vale cero en ejecucion
Private DesTime_Inset As Long

Private TWIPX As Long
Attribute TWIPX.VB_VarMemberFlags = "400"
Attribute TWIPX.VB_VarDescription = "Returns number of TWIPs equivalent to one Pixel on the user screen in the X dimension."
Private TWIPY As Long
Attribute TWIPY.VB_VarMemberFlags = "400"
Attribute TWIPY.VB_VarDescription = "Returns number of TWIPs equivalent to one Pixel on the user screen in the Y dimension."

'  Orientación
Public Enum Orientation_types
    HORZ = 0
    VERT = 1
End Enum

Private SenSIZE As Long
Private IndSIZE As Long

Private SplitLimitLO As Long
Private SplitLimitHI As Long

Private IndInset As Long

' Valores Predeterminados
Private Const m_def_Last_PerCent As Integer = 15
Private Const m_def_SplitPerCent As Integer = 50 ' default split %
Private Const m_def_Orientation As Byte = HORZ  ' default
Private Const m_def_SenseWidth As Integer = 4 ' default width of Slider sense region
Private Const m_def_Ind_Width As Integer = 2 ' width of Slider indicator
Private Const m_def_Ind_Inset As Integer = 4 ' inset value for Slider indicator
Private Const m_def_Limit_LO As Integer = 10 ' Slider travel LO-limit%
Private Const m_def_Limit_HI As Integer = 90 ' Slider travel HI-limit%
Private Const m_def_SPL_SwappedCtls = 0 ' = False

Private m_BackColor As OLE_COLOR

Private Const LO As String = "LO"
Private Const HI As String = "HI"

' variables de miembro
Private m_SPL_SwappedCtls As Long
Private m_SPL_Hide As Boolean
Private m_SPL_HideSUP As Boolean
Private m_SplitPerCent As Single
Private m_LastPerCent  As Single
Private m_Orientation As Long
'
Private m_Limit_LO As Long
Private m_Limit_HI As Long
'
Private m_SenseWidth As Long
Private m_Ind_Width As Long
Private m_Ind_Inset As Long
'para q no se vea como se resizean ventanas
''
''Public Sub LockWindow(FormHWnd As Long)
''    LockWindowUpdate FormHWnd
''End Sub

Private Sub imgSense_MouseDown(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
'============================================================================================
'  Comienza la operacion de arrastrar por el Raton
If m_SPL_Hide Then Exit Sub
    ' almacenamos la posicion del cursor
    xRefOffst = X
    yRefOffst = Y
       
    ' alineamos el indicador del Split encima del split
    If Not sliderMoving Then
        ' ¿q orientacion?
        If SPL_Orientation = HORZ Then
            ' muestra el split-indicator
            'picIndicator.Visible = True
            picIndicator.Refresh
            With imgSense
                picIndicator.Width = IndSIZE
                ' ajusta height del  indicador
                picIndicator.Height = .Height - (IndInset * 2)
                picIndicator.Left = .Left + SPLIND_CORRECT
                ' offset del indicator
                picIndicator.Top = .Top + IndInset
                ' ZOrder para el Indicador
                picIndicator.ZOrder 0
            End With
        Else
            ' muestra el indicador-split
            picIndicator.Visible = True
            picIndicator.Refresh
            With imgSense
                ' ajusta width del indicador
                picIndicator.Width = .Width - (IndInset * 2)
                picIndicator.Height = IndSIZE
                ' offset del indicador
                picIndicator.Left = .Left + IndInset
                picIndicator.Top = .Top + SPLIND_CORRECT
                'ZOrder
                picIndicator.ZOrder 0
            End With
        End If 'orientación
        sliderMoving = True
    End If
End Sub

Private Sub imgSense_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
'============================================================================================
'  se mueve por drag del ratón
If m_SPL_Hide Then Exit Sub
Dim sglPos As Single

    If sliderMoving Then
         '¿ orientacion¿
        If SPL_Orientation = HORZ Then
            ' posicion relativa al offset de Xreference
            sglPos = X + imgSense.Left - xRefOffst + SPLIND_CORRECT
            ' Limite LO -- test
            If sglPos < SplitLimitLO Then
               picIndicator.Left = SplitLimitLO
              ' Limite HI -- test
            ElseIf sglPos > (SplitLimitHI) Then
               picIndicator.Left = (SplitLimitHI)
            Else
              ' indicador q sigue al raton
               picIndicator.Left = sglPos
            End If
        Else
             ' calcula posicion relativa al offset Yreference
             sglPos = Y + imgSense.Top - yRefOffst + SPLIND_CORRECT
             ' Limite LO -- test (puede liarse un poco)
             If sglPos > SplitLimitLO Then
                picIndicator.Top = SplitLimitLO
                ' Limite HI -- test
             ElseIf sglPos < (SplitLimitHI) Then
                picIndicator.Top = (SplitLimitHI)
             Else
               ' indicador q sigue al raton
                picIndicator.Top = sglPos
             End If
        End If ' orientacion
    End If ' moving
End Sub

Private Sub imgSense_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'=========================================================================================
'  Paso Final de seleccion del splitter

    picIndicator.Visible = False
    picIndicator.Refresh
    
    If sliderMoving Then
         ' ¿orientacion? -- Calcula nuevo Split perCent
        If SPL_Orientation = HORZ Then
            ' Actualiza la  posicion del Split-Sense para cubrir el Split-Indicator
             imgSense.Left = picIndicator.Left - SPLIND_CORRECT
             m_SplitPerCent = (picIndicator.Left / UserControl.Width) * 100
        Else
            ' Actualiza la  posicion del Split-Sense para cubrir el Split-Indicator
             imgSense.Top = picIndicator.Top - SPLIND_CORRECT
             m_SplitPerCent = (1 - (picIndicator.Top / UserControl.Height)) * 100
        End If
    End If
    
    ' asegurarse de q sense esta arriba
    imgSense.ZOrder 0
    
    ' limpiamos flag
    sliderMoving = False
   
   ' rearrange controls hosted by this UserControl
   SizeControls
   
   ' supply new split info via event
   RaiseEvent NewSplit(CLng(m_SplitPerCent))
End Sub

Private Sub SizeControls()
Attribute SizeControls.VB_Description = "Moves and sizes the two controls hosted by the splitter. These controls -must- have their TAG properties set to ""HI"" or ""LO"" respectively."
'================================== Arrange and size controls on UserControl
Dim CTL As Control

    ' enforce scalemode setting
    UserControl.ScaleMode = vbTwips
    On Error Resume Next
   
   ' only execute if percent setting has changed
   If (LastPercent <> m_SplitPerCent) Then
        ' if no controls to size -- just exit
        If UserControl.ContainedControls.Count = 0 Then Exit Sub
        ' find controls marked as HI or LO in (Tag property) and size
        For Each CTL In UserControl.ContainedControls
            If UCase$(CTL.Tag) = LO Then '===================== LO
                ' size the left or bottom pane
                With CTL
                    ' pay attention to orientation
                    If SPL_Orientation = HORZ Then
                        .Left = DesTime_Inset
                        .Top = DesTime_Inset
                        .Width = IIf(imgSense.Left - (DesTime_Inset * 2) < 0, 0, imgSense.Left - (DesTime_Inset * 2))
                        .Height = IIf(UserControl.Height - (DesTime_Inset * 2) < 0, 0, UserControl.Height - (DesTime_Inset * 2))
                    Else
                        .Left = DesTime_Inset
                        .Top = imgSense.Top + imgSense.Height + DesTime_Inset
                        .Width = IIf(UserControl.Width - .Left - (DesTime_Inset * 1) < 0, 0, UserControl.Width - .Left - (DesTime_Inset * 1))
                        
                        .Height = IIf(Abs(UserControl.Height - .Top) - (DesTime_Inset * 1) < 0, 0, Abs(UserControl.Height - .Top) - (DesTime_Inset * 1))
                    End If ' SPL_Orientation
                   .ZOrder 1
                   imgSense.ZOrder 0
                End With
            ElseIf UCase$(CTL.Tag) = HI Then ' ================= HI
                ' size the right or top pane
                With CTL
                    ' pay attention to orientation
                    If SPL_Orientation = HORZ Then
                        .Left = imgSense.Left + imgSense.Width + DesTime_Inset
                        .Top = DesTime_Inset
                        .Width = IIf(UserControl.Width - .Left - (DesTime_Inset * 1) < 0, 0, UserControl.Width - .Left - (DesTime_Inset * 1))
                        .Height = IIf(UserControl.Height - (DesTime_Inset * 2) < 0, 0, UserControl.Height - (DesTime_Inset * 2))
                    Else
                        .Left = DesTime_Inset
                        .Top = DesTime_Inset
                        .Width = IIf(UserControl.Width - (DesTime_Inset * 2) < 0, 0, UserControl.Width - (DesTime_Inset * 2))
                        .Height = IIf(imgSense.Top - (DesTime_Inset * 2) < 0, 0, imgSense.Top - (DesTime_Inset * 2))
                    End If 'SPL_Orientation
                   .ZOrder 1
                   imgSense.ZOrder 0
                End With
            End If ' ucase$(CTL.Tag)
        Next
    End If ' lastpercent
    
    ' update detector
    LastPercent = m_SplitPerCent
    ' make sure sense control is at top of Zorder
    imgSense.ZOrder 0
    RaiseEvent Sized 'MaRio.-
End Sub

Private Sub SetView()
On Error Resume Next
'===================
'  Here, we set up the basic items that control splitter operation
    
    ' init Range LIMITS -- pay attention to orientation
    If SPL_Orientation = HORZ Then
        SplitLimitLO = (m_Limit_LO / 100) * UserControl.Width
        SplitLimitHI = (m_Limit_HI / 100) * UserControl.Width
        ' SET the E-W mousepointer
        imgSense.MousePointer = 9
        ' SET Starting SPLITTER SENSE and WIDTH variables
        SenSIZE = m_SenseWidth * TWIPX
        IndSIZE = m_Ind_Width * TWIPX
        ' init INDICATOR INSET variable
        IndInset = m_Ind_Inset * TWIPX
    Else
        ' must recompute for proper limits
        SplitLimitLO = (1 - (m_Limit_LO / 100)) * UserControl.Height
        SplitLimitHI = (1 - (m_Limit_HI / 100)) * UserControl.Height
        ' SET the N-S mousepointer
        imgSense.MousePointer = 7
        ' SET Starting SPLITTER SENSE and WIDTH variables
        SenSIZE = m_SenseWidth * TWIPY
        IndSIZE = m_Ind_Width * TWIPY
        ' init INDICATOR INSET variable
        IndInset = m_Ind_Inset * TWIPY
    End If
   
    ' calculate correction value SENSE VS. INDICATOR
    SPLIND_CORRECT = (Abs(SenSIZE - IndSIZE) / 2)
    
    ' pre-align Slider Sense and Indicator  -- pay attention to orientation
    If SPL_Orientation = HORZ Then
    '======== HORIZ WINDOW
        picIndicator.Left = UserControl.Width * (m_SplitPerCent / 100)
        picIndicator.Top = IndInset
        picIndicator.Width = IndSIZE
        picIndicator.Height = UserControl.Height - (IndInset * 2)
        '
        imgSense.Left = picIndicator.Left - SPLIND_CORRECT
        imgSense.Top = 0
        imgSense.Width = SenSIZE
        imgSense.Height = UserControl.Height
    Else
    '======== VERT WINDOW
        picIndicator.Left = IndInset
        picIndicator.Height = IndSIZE
        picIndicator.Top = UserControl.Height * (1 - (m_SplitPerCent / 100))
        picIndicator.Width = UserControl.Width - (IndInset * 2)
        picIndicator.Refresh
        '
        imgSense.Left = 0
        imgSense.Top = picIndicator.Top - SPLIND_CORRECT
        imgSense.Width = UserControl.Width
        imgSense.Height = SenSIZE
        imgSense.Refresh
    End If
    
    ' reposition and size Splitter Hosted controls
    SizeControls
End Sub



Private Sub UserControl_Resize()
'==============================
    ' set var to value guaranteed to cause an update at "resize controls" sub
    LastPercent = 110
    ' start at a known state
    DesTime_Inset = 0
    
    ' Are we are in design_time mode?
    If Not UserControl.Ambient.UserMode Then
        ' YES
        UserControl.BorderStyle = 0
            ' draw borders around control -- helps visualize at designtime
            With UserControl
                .AutoRedraw = True
                .Cls
                UserControl.Line (0, 0)-(.Width, .Height), , B
            End With
            
        ' you can set this var to zero if it annoys you, Don't make >10!
        ' used to make hosted controls somewhat smaller at design_time
        ' easier to grab control on Form this way
        DesTime_Inset = 10 * TWIPX
        
        ' show the border on Slider Sense-window
        imgSense.BorderStyle = 1
        ' make the Slider indicator visible
        picIndicator.Visible = True
        ' force Slider indicator to top of zOrder
        picIndicator.ZOrder 0
        UserControl.Refresh
    End If
    
    ' force the view into alignment (calls resize controls)
    SetView
End Sub

Private Sub UserControl_Initialize()
'--------------------------------------
    UserControl.ScaleMode = vbTwips
    TWIPX = Screen.TwipsPerPixelX
    TWIPY = Screen.TwipsPerPixelY
End Sub

'  Initialize default Properties for User Control when dropped onto a Form
Private Sub UserControl_InitProperties()
'====================================
    m_SplitPerCent = m_def_SplitPerCent
    m_Orientation = m_def_Orientation
    m_SenseWidth = m_def_SenseWidth
    m_Limit_LO = m_def_Limit_LO
    m_Limit_HI = m_def_Limit_HI
    m_Ind_Width = m_def_Ind_Width
    m_Ind_Inset = m_def_Ind_Inset
    m_SPL_SwappedCtls = m_def_SPL_SwappedCtls
End Sub

'  Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'=============================================================

    Call PropBag.WriteProperty("BackClr", UserControl.BackColor, UserControl.BackColor)
    Call PropBag.WriteProperty("SplitHide", m_SPL_Hide, False)
    Call PropBag.WriteProperty("SplitHideSUP", m_SPL_HideSUP, False)
    Call PropBag.WriteProperty("SplitPerCent", m_SplitPerCent, m_def_SplitPerCent)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("SenseWidth", m_SenseWidth, m_def_SenseWidth)
    Call PropBag.WriteProperty("Limit_LO", m_Limit_LO, m_def_Limit_LO)
    Call PropBag.WriteProperty("Limit_HI", m_Limit_HI, m_def_Limit_HI)
    Call PropBag.WriteProperty("Ind_Width", m_Ind_Width, m_def_Ind_Width)
    Call PropBag.WriteProperty("Ind_Inset", m_Ind_Inset, m_def_Ind_Inset)
    Call PropBag.WriteProperty("Ind_Color", picIndicator.BackColor, &HFF0000)
    Call PropBag.WriteProperty("SPL_SwappedCtls", m_SPL_SwappedCtls, m_def_SPL_SwappedCtls)
End Sub

'  Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'============================================================
    UserControl.BackColor = PropBag.ReadProperty("BackClr", UserControl.BackColor)
    m_SplitPerCent = PropBag.ReadProperty("SplitPerCent", m_def_SplitPerCent)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_SenseWidth = PropBag.ReadProperty("SenseWidth", m_def_SenseWidth)
    m_Limit_LO = PropBag.ReadProperty("Limit_LO", m_def_Limit_LO)
    m_Limit_HI = PropBag.ReadProperty("Limit_HI", m_def_Limit_HI)
    m_Ind_Width = PropBag.ReadProperty("Ind_Width", m_def_Ind_Width)
    m_Ind_Inset = PropBag.ReadProperty("Ind_Inset", m_def_Ind_Inset)
    picIndicator.BackColor = PropBag.ReadProperty("Ind_Color", &H8000000D)
    m_SPL_SwappedCtls = PropBag.ReadProperty("SPL_SwappedCtls", m_def_SPL_SwappedCtls)
        If UserControl.Ambient.UserMode Then
            imgSense.BorderStyle = 0
        Else
            imgSense.BorderStyle = 1
        End If
    UserControl_Resize
End Sub


'==============================================================  PerCent Split Property
Public Property Get SPL_PerCent() As Long
Attribute SPL_PerCent.VB_Description = "Returns/sets the -split- of the UserControl display area in per-cent, relative to the LO designated viewport.  At 100%, the LO viewport fills the entire UserControl display."
Attribute SPL_PerCent.VB_UserMemId = 0
    SPL_PerCent = m_SplitPerCent
End Property
Public Property Let SPL_PerCent(ByVal New_SplitPerCent As Long)
    ' only allow update when value is new
    If New_SplitPerCent <> m_SplitPerCent Then
        m_SplitPerCent = New_SplitPerCent
        ' keep within limit bounds
        If m_SplitPerCent > m_Limit_HI Then m_SplitPerCent = m_Limit_HI
        If m_SplitPerCent < m_Limit_LO Then m_SplitPerCent = m_Limit_LO
        PropertyChanged "SplitPerCent"
        UserControl_Resize
    End If
End Property

Public Property Get SPL_Hide() As Boolean
    
    If m_SplitPerCent = 0 Then
        SPL_Hide = True
        m_SPL_Hide = True
    End If
End Property

Public Property Get BackClr() As OLE_COLOR
    '
    BackClr = UserControl.BackColor
End Property
Public Property Let BackClr(v As OLE_COLOR)
    '
    m_BackColor = v
    UserControl.BackColor = v
    PropertyChanged "BackClr"
    'UserControl_Resize
End Property


Public Property Get hWnd() As Long
     hWnd = UserControl.hWnd
End Property

Public Property Let SPL_Hide(v As Boolean)
    If v Then
        If m_SplitPerCent <> 0 And m_SplitPerCent <> 100 Then
            m_LastPerCent = m_SplitPerCent
        End If
        m_SplitPerCent = 0
        m_SPL_Hide = True
        m_SPL_HideSUP = False
    Else

       If m_SplitPerCent <> 0 And m_SplitPerCent <> 100 Then
            m_LastPerCent = m_SplitPerCent
       'Else
       'if m_LastPerCent m_LastPerCent = m_def_Last_PerCent
       End If

        
        
        m_SplitPerCent = m_LastPerCent
        m_SPL_Hide = False
        m_SPL_HideSUP = False
    End If
    PropertyChanged "SplitPerCent"
    PropertyChanged "SplitHide"
    UserControl_Resize
End Property

Public Property Get SPL_HideSUP() As Boolean
    
    If m_SplitPerCent = 100 Then
        SPL_HideSUP = True
        m_SPL_HideSUP = True
    End If
End Property
Public Property Let SPL_HideSUP(v As Boolean)
    If v Then
        If m_SplitPerCent <> 0 And m_SplitPerCent <> 100 Then
            m_LastPerCent = m_SplitPerCent
        End If
        m_SplitPerCent = 100
        m_SPL_HideSUP = True
        m_SPL_Hide = False
    Else
       If m_SplitPerCent <> 0 And m_SplitPerCent <> 100 Then
            m_LastPerCent = m_SplitPerCent
       'Else
       ' m_LastPerCent = m_def_Last_PerCent
       End If

        m_SplitPerCent = m_LastPerCent
        m_SPL_HideSUP = False
        m_SPL_Hide = False
    End If
    PropertyChanged "SplitPerCent"
    PropertyChanged "SplitHideSUP"
    UserControl_Resize
End Property




'===============================================================  Orientation Property
Public Property Get SPL_Orientation() As Orientation_types
Attribute SPL_Orientation.VB_Description = "Returns/sets whether Splitter is Horizontal or Vertical in orientation. Slider travel is left-right for Horizontal and up-down for Vertical orientation."
    SPL_Orientation = m_Orientation
End Property
Public Property Let SPL_Orientation(ByVal New_Orientation As Orientation_types)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
    UserControl_Resize
End Property


'=============================================================  LO Split-Limit Property
Public Property Get SPL_Limit_LO() As Long
Attribute SPL_Limit_LO.VB_Description = "Returns/sets Lower-Limit of Slider travel in per-cent."
    SPL_Limit_LO = m_Limit_LO
End Property
Public Property Let SPL_Limit_LO(ByVal New_Limit_LO As Long)
    m_Limit_LO = New_Limit_LO
    ' don't allow less than 10%
    If m_Limit_LO < 10 Then m_Limit_LO = 10
    PropertyChanged "Limit_LO"
    SplitLimitLO = (m_Limit_LO / 100) * UserControl.Width
End Property


'============================================================  HI Split-Limit Property
Public Property Get SPL_Limit_HI() As Long
Attribute SPL_Limit_HI.VB_Description = "Returns/sets Upper-Limit of Slider travel in per-cent."
    SPL_Limit_HI = m_Limit_HI
End Property
Public Property Let SPL_Limit_HI(ByVal New_Limit_HI As Long)
    m_Limit_HI = New_Limit_HI
    ' don't allow greater than 90%
    If m_Limit_HI > 90 Then m_Limit_HI = 90
    PropertyChanged "Limit_HI"
    SplitLimitHI = (m_Limit_HI / 100) * UserControl.Width
End Property


'==========================================================  Slider Sense-Width Property
Public Property Get SPL_SenseWidth() As Long
Attribute SPL_SenseWidth.VB_Description = "Returns/Sets width of Slider sense-region in pixels."
    SPL_SenseWidth = m_SenseWidth
End Property
Public Property Let SPL_SenseWidth(ByVal New_SenseWidth As Long)
    m_SenseWidth = New_SenseWidth
    PropertyChanged "SenseWidth"
    
    SenSIZE = m_SenseWidth
    
    ' keep Slider indicator-width within limits relative to Slider sense-width
    If m_Ind_Width > (m_SenseWidth - 2) Then m_Ind_Width = (m_SenseWidth - 2)
    If m_Ind_Width < 1 Then m_Ind_Width = 1
    IndSIZE = m_Ind_Width
    
    ' calculate correction value SENSE VS. INDICATOR
    SPLIND_CORRECT = (Abs(SenSIZE - IndSIZE) / 2)
    UserControl_Resize
End Property


'======================================================  Slider Indicator-Width Property
Public Property Get SPL_Ind_Width() As Long
Attribute SPL_Ind_Width.VB_Description = "Returns/sets Width of Slider indicator in pixels."
    SPL_Ind_Width = m_Ind_Width
End Property
Public Property Let SPL_Ind_Width(ByVal New_Ind_Width As Long)
    m_Ind_Width = New_Ind_Width
    PropertyChanged "Ind_Width"
    
    ' keep Slider Indicator-width within limits relative to Slider Sense-width
    If m_Ind_Width > (m_SenseWidth - 2) Then m_Ind_Width = (m_SenseWidth - 2)
    If m_Ind_Width < 1 Then m_Ind_Width = 1
    IndSIZE = m_Ind_Width
    
        ' calculate correction value SENSE VS. INDICATOR
    SPLIND_CORRECT = (Abs(SenSIZE - IndSIZE) / 2)
    UserControl_Resize
End Property


'======================================================  Slider Indicator-Inset Property
Public Property Get SPL_Ind_Inset() As Long
Attribute SPL_Ind_Inset.VB_Description = "Returns/sets Slider indicator inset or margin in pixels.  A value of zero equals full length indication."
    SPL_Ind_Inset = m_Ind_Inset
End Property
Public Property Let SPL_Ind_Inset(ByVal New_Ind_Inset As Long)
    m_Ind_Inset = New_Ind_Inset
    PropertyChanged "Ind_Inset"
    UserControl_Resize
End Property


'=======================================================  Slider Indicator-Color Property
'MappingInfo=picIndicator,picIndicator,-1,BackColor
Public Property Get SPL_Ind_Color() As OLE_COLOR
Attribute SPL_Ind_Color.VB_Description = "Returns/sets Slider indicator color."
    SPL_Ind_Color = picIndicator.BackColor
End Property
Public Property Let SPL_Ind_Color(ByVal New_Ind_Color As OLE_COLOR)
    picIndicator.BackColor() = New_Ind_Color
    PropertyChanged "Ind_Color"
    UserControl_Resize
End Property


'===================================================  Swapped Hosted-Controls Property
Public Property Get SPL_SwappedCtls() As Boolean
Attribute SPL_SwappedCtls.VB_Description = "Returns/sets whether hosted controls are swapped regarding LO-HI Tag-property assignments."
    SPL_SwappedCtls = m_SPL_SwappedCtls
End Property
Public Property Let SPL_SwappedCtls(ByVal New_SPL_SwappedCtls As Boolean)
    m_SPL_SwappedCtls = New_SPL_SwappedCtls
    PropertyChanged "SPL_SwappedCtls"
    
    Dim CTL As Control
    ' swap control LO-HI assignments
    For Each CTL In UserControl.ContainedControls
        If UCase$(CTL.Tag) = LO Then
            CTL.Tag = HI
        ElseIf UCase$(CTL.Tag) = HI Then
            CTL.Tag = LO
        End If
    Next
    UserControl_Resize
End Property

