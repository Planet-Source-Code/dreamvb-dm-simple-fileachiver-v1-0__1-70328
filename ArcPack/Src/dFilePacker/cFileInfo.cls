VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Filename As String
Private m_FileSize As Long

Public Property Get Filename() As String
    Filename = m_Filename
End Property

Public Property Let Filename(ByVal vNewValue As String)
    m_Filename = vNewValue
End Property

Public Property Get FileLength() As String
    FileLength = m_FileSize
End Property

Public Property Let FileLength(ByVal vNewValue As String)
    m_FileSize = vNewValue
End Property

