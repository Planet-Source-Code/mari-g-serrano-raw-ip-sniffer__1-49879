VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
' Mensajes de Windows :(25/07/02)
Public Enum WM
    WM_NULL = &H0
    WM_CREATE = &H1
    WM_DESTROY = &H2
    WM_MOVE = &H3
    WM_SIZE = &H5
    WM_ACTIVATE = &H6
    WM_SETFOCUS = &H7
    WM_KILLFOCUS = &H8
    WM_ENABLE = &HA
    WM_SETREDRAW = &HB
    WM_SETTEXT = &HC
    WM_GETTEXT = &HD
    WM_GETTEXTLENGTH = &HE
    WM_PAINT = &HF
    WM_CLOSE = &H10
    WM_QUERYENDSESSION = &H11
    WM_QUIT = &H12
    WM_QUERYOPEN = &H13
    WM_ERASEBKGND = &H14
    WM_SYSCOLORCHANGE = &H15
    WM_ENDSESSION = &H16
    WM_SHOWWINDOW = &H18
    WM_WININICHANGE = &H1A
    WM_DEVMODECHANGE = &H1B
    WM_ACTIVATEAPP = &H1C
    WM_FONTCHANGE = &H1D
    WM_TIMECHANGE = &H1E
    WM_CANCELMODE = &H1F
    WM_SETCURSOR = &H20
    WM_MOUSEACTIVATE = &H21
    WM_CHILDACTIVATE = &H22
    WM_QUEUESYNC = &H23
    WM_GETMINMAXINFO = &H24
    WM_PAINTICON = &H26
    WM_ICONERASEBKGND = &H27
    WM_NEXTDLGCTL = &H28
    WM_SPOOLERSTATUS = &H2A
    WM_DRAWITEM = &H2B
    WM_MEASUREITEM = &H2C
    WM_DELETEITEM = &H2D
    WM_VKEYTOITEM = &H2E
    WM_CHARTOITEM = &H2F
    WM_SETFONT = &H30
    WM_GETFONT = &H31
    WM_SETHOTKEY = &H32
    WM_GETHOTKEY = &H33
    WM_QUERYDRAGICON = &H37
    WM_COMPAREITEM = &H39
    WM_COMPACTING = &H41
    WM_OTHERWINDOWCREATED = &H42
    WM_OTHERWINDOWDESTROYED = &H43
    WM_COMMNOTIFY = &H44
    WM_WINDOWPOSCHANGING = &H46
    WM_WINDOWPOSCHANGED = &H47
    WM_POWER = &H48
    WM_COPYDATA = &H4A
    WM_CANCELJOURNAL = &H4B
    WM_NCCREATE = &H81
    WM_NCDESTROY = &H82
    WM_NCCALCSIZE = &H83
    WM_NCHITTEST = &H84
    WM_NCPAINT = &H85
    WM_NCACTIVATE = &H86
    WM_GETDLGCODE = &H87
    WM_NCMOUSEMOVE = &HA0
    WM_NCLBUTTONDOWN = &HA1
    WM_NCLBUTTONUP = &HA2
    WM_NCLBUTTONDBLCLK = &HA3
    WM_NCRBUTTONDOWN = &HA4
    WM_NCRBUTTONUP = &HA5
    WM_NCRBUTTONDBLCLK = &HA6
    WM_NCMBUTTONDOWN = &HA7
    WM_NCMBUTTONUP = &HA8
    WM_NCMBUTTONDBLCLK = &HA9
    WM_KEYFIRST = &H100
    WM_KEYDOWN = &H100
    WM_KEYUP = &H101
    WM_CHAR = &H102
    WM_DEADCHAR = &H103
    WM_SYSKEYDOWN = &H104
    WM_SYSKEYUP = &H105
    WM_SYSCHAR = &H106
    WM_SYSDEADCHAR = &H107
    WM_KEYLAST = &H108
    WM_INITDIALOG = &H110
    WM_COMMAND = &H111
    WM_SYSCOMMAND = &H112
    WM_TIMER = &H113
    WM_HSCROLL = &H114
    WM_VSCROLL = &H115
    WM_INITMENU = &H116
    WM_INITMENUPOPUP = &H117
    WM_MENUSELECT = &H11F
    WM_MENUCHAR = &H120
    WM_ENTERIDLE = &H121
    WM_CTLCOLORMSGBOX = &H132
    WM_CTLCOLOREDIT = &H133
    WM_CTLCOLORLISTBOX = &H134
    WM_CTLCOLORBTN = &H135
    WM_CTLCOLORDLG = &H136
    WM_CTLCOLORSCROLLBAR = &H137
    WM_CTLCOLORSTATIC = &H138
    WM_MOUSEFIRST = &H200
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_LBUTTONDBLCLK = &H203
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
    WM_RBUTTONDBLCLK = &H206
    WM_MBUTTONDOWN = &H207
    WM_MBUTTONUP = &H208
    WM_MBUTTONDBLCLK = &H209
    WM_MOUSELAST = &H209
    WM_PARENTNOTIFY = &H210
    WM_ENTERMENULOOP = &H211
    WM_EXITMENULOOP = &H212
    WM_MDICREATE = &H220
    WM_MDIDESTROY = &H221
    WM_MDIACTIVATE = &H222
    WM_MDIRESTORE = &H223
    WM_MDINEXT = &H224
    WM_MDIMAXIMIZE = &H225
    WM_MDITILE = &H226
    WM_MDICASCADE = &H227
    WM_MDIICONARRANGE = &H228
    WM_MDIGETACTIVE = &H229
    WM_MDISETMENU = &H230
    WM_DROPFILES = &H233
    WM_MDIREFRESHMENU = &H234
    WM_CUT = &H300
    WM_COPY = &H301
    WM_PASTE = &H302
    WM_CLEAR = &H303
    WM_UNDO = &H304
    WM_RENDERFORMAT = &H305
    WM_RENDERALLFORMATS = &H306
    WM_DESTROYCLIPBOARD = &H307
    WM_DRAWCLIPBOARD = &H308
    WM_PAINTCLIPBOARD = &H309
    WM_VSCROLLCLIPBOARD = &H30A
    WM_SIZECLIPBOARD = &H30B
    WM_ASKCBFORMATNAME = &H30C
    WM_CHANGECBCHAIN = &H30D
    WM_HSCROLLCLIPBOARD = &H30E
    WM_QUERYNEWPALETTE = &H30F
    WM_PALETTEISCHANGING = &H310
    WM_PALETTECHANGED = &H311
    WM_HOTKEY = &H312
    WM_PENWINFIRST = &H380
    WM_PENWINLAST = &H38F
    WM_USER = &H400
    WM_DDE_FIRST = &H3E0
    WM_DDE_INITIATE = (WM_DDE_FIRST)
    WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
    WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
    WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
    WM_DDE_ACK = (WM_DDE_FIRST + 4)
    WM_DDE_DATA = (WM_DDE_FIRST + 5)
    WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
    WM_DDE_POKE = (WM_DDE_FIRST + 7)
    WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
    WM_DDE_LAST = (WM_DDE_FIRST + 8)
    WM_CONVERTREQUESTEX = &H108
    WM_IME_STARTCOMPOSITION = &H10D
    WM_IME_ENDCOMPOSITION = &H10E
    WM_IME_COMPOSITION = &H10F
    WM_IME_KEYLAST = &H10F
    WM_IME_SETCONTEXT = &H281
    WM_IME_NOTIFY = &H282
    WM_IME_CONTROL = &H283
    WM_IME_COMPOSITIONFULL = &H284
    WM_IME_SELECT = &H285
    WM_IME_CHAR = &H286
    WM_IME_KEYDOWN = &H290
    WM_IME_KEYUP = &H291
    WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
    WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
    WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)
    WM_PSD_PAGESETUPDLG = (WM_USER)
    WM_PSD_FULLPAGERECT = (WM_USER + 1)
    WM_PSD_MINMARGINRECT = (WM_USER + 2)
    WM_PSD_MARGINRECT = (WM_USER + 3)
    WM_PSD_GREEKTEXTRECT = (WM_USER + 4)
    WM_PSD_ENVSTAMPRECT = (WM_USER + 5)
    WM_PSD_YAFULLPAGERECT = (WM_USER + 6)
    WM_MOUSEWHEEL = &H20A
    WINSOCKMSG = (WM_USER + 1)
    
End Enum





Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function htonl Lib "ws2_32.dll" (ByVal HostLong As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC       As Long = (-4)      'Establece la direccion del WndProc con GetWindowLong/SetWindowLong
Private Const BUF_TOP           As Long = 511       'Offset maximo en buffer OpCode. necesita 136 + (# Messages * 11)
Private Const OFFSET_BEFORE     As Long = 3         'Offset dentro de los bytes de opcode para el procesamiento Antes del por Defecto
Private Const OFFSET_AFTER      As Long = 65        'Offset dentro de los bytes de opcode para el procesamiento Despues del por Defecto
Private Const CODE_RETURN       As Long = &H10C2C9  'por defecto...
Private Const OPCODE_CMP_EAX    As String = "3D"    'Opcode para cmp eax,????????
Private Const OPCODE_JE         As String = "0F84"  'Opcode para je con un offset relativo de 4 bytes.
'Codigo Maquina:
Private Const WNDPROC_ALL       As String = "558BEC83C4FCFF7514FF7510FF750CFF7508E8wnd_proc8945FCFF7514FF7510FF750CFF75088D45FC50B8ptrOwner8BC88B0950FF511C8B45FCC9C21000"
Private Const WNDPROC_FILTERED  As String = "558BEC83C4F8EB6AC745FC000000008D45FC50C745F8000000008D45F850B8ptrOwner8BC88B0950FF5120837DF800753AFF7514FF7510FF750CFF7508E8wnd_procC9C21000E8wnd_proc8945FCFF7514FF7510FF750CFF75088D45FC50B8ptrOwner8BC88B0950FF511C8B45FCC9C210008B450CFF7514FF751050FF7508"
Private Const MSG_UNHANDLED     As String = "E8wnd_procC9C21000"

Private Type tCode
  Buf(0 To BUF_TOP) As Byte       ' Buffer de Opcodes
End Type

Private Type tCodeBuf
  Code              As tCode      ' opcodes para WndProc
  nBuf              As Long       ' indice de Opcode en el Buffer
End Type

Private All         As Boolean    ' Todos los mensajes?
Private Running     As Boolean    '
Private hWnd        As Long       '
Private WndProcPrev As Long       ' direccion del WndProc
Private pCode       As Long       ' Puntero al buffer de Opcodes de WndProc
Private CodeBuf     As tCodeBuf   ' buffer de Opcodes


Public Sub AddMsg(MsgNum As WM, Optional Antes As Boolean = False)
  Debug.Assert (Running = False)
  
  With CodeBuf
    If .nBuf = 0 Then AddCode (WNDPROC_FILTERED)
    AddCode (OPCODE_CMP_EAX & Hex8(htonl(MsgNum)))
    AddCode (OPCODE_JE & Hex8(htonl(Not (.nBuf - IIf(Antes, OFFSET_BEFORE, OFFSET_AFTER)))))
  End With
End Sub

Public Sub SubClass(hWndSub As Long, Owner As iSubCls, Optional AllMsgs As Boolean = False)
  Dim pOwner  As Long
  Dim nPos    As Long
  
  All = AllMsgs
  
  With CodeBuf
    Debug.Assert (Running = False)
    Debug.Assert (IsWindow(hWndSub))
    Debug.Assert (Not All And .nBuf > 0) Or _
                 (All And .nBuf = 0)
    hWnd = hWndSub
    WndProcPrev = GetWindowLong(hWnd, GWL_WNDPROC)
    
    pOwner = ObjPtr(Owner) 'puntero a frmMain
    pCode = VarPtr(.Code.Buf(0)) 'Direccion de nuestro WndProc
    
    If AllMsgs Then
    
      Call AddCode(WNDPROC_ALL)
      Call PatchOffset(19)                        'Parchea la llamada a WndProcPrev
      Call PatchValue(43, pOwner)
    Else
    
      Call PatchValue(31, pOwner)
      Call PatchOffset(62)                        'Parchea la llamada anterior a WndProcPrev
      Call PatchOffset(71)                        'Parchea la llamada posterior a WndProcPrev
      Call PatchValue(95, pOwner)
      
      nPos = .nBuf + 1                                  'guarda la posicion del buf
      Call AddCode(MSG_UNHANDLED)                       'Añade la llamada a WndProcPrev
      Call PatchOffset(nPos)                            'parchea  WndProcPrev
    End If
  End With
  
  'Call Crash
  'que windows llame al proc generado en ejecucion (un array)
  Call SetWindowLong(hWnd, GWL_WNDPROC, pCode)
  Running = True
End Sub

Public Sub UnSubclass()
  If Running Then
    If All Then
      Call PatchValue(23, CODE_RETURN)
    Else
      CodeBuf.Code.Buf(7) = &H29
      Call PatchValue(75, CODE_RETURN)
    End If
    
    Call SetWindowLong(hWnd, GWL_WNDPROC, WndProcPrev)  'Restaura el WndProc
    CodeBuf.nBuf = 0                                    'limpia el buffer de Opcodes
    Running = False
  End If
End Sub

Private Sub Class_Terminate()
  If Running Then UnSubclass
End Sub

Private Sub AddCode(sOps As String)
  Dim i As Long
  Dim j As Long
  
  With CodeBuf
    j = Len(sOps)
    Debug.Assert (.nBuf + (j \ 2) <= BUF_TOP)
    
    For i = 1 To j Step 2
      
      .Code.Buf(.nBuf) = Val("&H" & Mid$(sOps, i, 2))
      .nBuf = .nBuf + 1
    Next i
  End With
End Sub

Private Function Hex8(lValue As Long) As String
  Dim s As String
  s = Hex$(lValue)
  Hex8 = String$(8 - Len(s), "0") & s
End Function

'Parchea el buffer en la direccion offset con el valor pasado
Private Sub PatchValue(nOffset As Long, nValue As Long)
  Call CopyMemory(ByVal (pCode + nOffset), nValue, 4)
End Sub

'Parchea el buffer cor el offset relativo al WndProc previo
Private Sub PatchOffset(nOffset As Long)
  Call CopyMemory(ByVal (pCode + nOffset), WndProcPrev - pCode - nOffset - 4, 4)
End Sub

'Debug:
'Private Sub Crash()
' Call CopyMemory(ByVal 0, &HFFFFFFFF, 1)
'End Sub
