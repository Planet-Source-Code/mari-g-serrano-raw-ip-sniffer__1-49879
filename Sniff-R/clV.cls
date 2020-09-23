VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public mProtocol As String
Public mSrcAddr  As String
Public mDestAddr As String
Public mSrcPort  As String
Public mDestPort As String
Public mSize     As String
Public mData     As String
Public mHora     As String
Public Sub add(Proto As String, SrcAddr As String, SrcPort As String, DestAddr As String, DestPort As String, Size As String, Data As String, Hora As String)
    mProtocol = Proto
    mSrcAddr = SrcAddr
    mDestAddr = DestAddr
    mSrcPort = SrcPort
    mDestPort = DestPort
    mSize = Size
    mData = Data
    mHora = Hora
End Sub
