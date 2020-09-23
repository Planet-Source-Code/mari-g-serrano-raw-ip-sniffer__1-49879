VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iSubCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Interface para "Eventos"

Option Explicit

Public Sub Despues(ByRef lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  'lReturn  - Obtiene valor de lReturn como mensaje/situacion/Windows necesita
  'hWnd     - Handle de la ventana
  'uMsg     - Nº de Mensaje
  'wParam   - datos relaccionados
  'lParam   - datos relaccionados
End Sub

Public Sub Antes(ByRef lHandled As Long, ByRef lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  'lHandled - si lHandled<>0 el WndProc por defecto no procesará este mensaje
  'lReturn  - establece el valor de lReturn como mensaje/situacion/Windows necesita
  'hWnd     - Handle de la ventana
  'uMsg     - Nº de Mensaje
  'wParam   - datos relaccionados
  'lParam   - datos relaccionados
End Sub
