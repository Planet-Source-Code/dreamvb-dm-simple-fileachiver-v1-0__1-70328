VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dFileArc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'My Simple File Achiver
'By DreamVB
'Version 1

Option Explicit

Private Type ARCHEADER
    ID As String * 3
    Version As Byte
    NOF As Long
    HeadSize As Integer
    FileSize As Long
End Type

Private Type FILEINFO
    Filename As String
    FileLength As Long
End Type

Private m_Filename As String
'Achive header
Private m_ArcHeader As ARCHEADER
'File info.
Private fInfo As FILEINFO
'Tells us if the file is open
Private m_IsOpen As Boolean
'File Pointer.
Private dFilePtr As Integer

Public Property Get Signature() As String
    Signature = m_ArcHeader.ID
End Property

Public Property Get FileCount() As Long
    FileCount = m_ArcHeader.NOF
End Property

Public Property Get Version() As Byte
    Version = m_ArcHeader.Version
End Property

Public Property Let Version(ByVal NewVer As Byte)
    m_ArcHeader.Version = NewVer
End Property

Public Property Get Filename() As String
    Filename = m_Filename
End Property

Public Property Let Filename(ByVal NewFile As String)
    m_Filename = NewFile
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = m_IsOpen
End Property

Public Property Get FileLength() As Long
    FileLength = m_ArcHeader.FileSize
End Property

Public Property Let IsOpen(ByVal vNewValue As Boolean)
    m_IsOpen = vNewValue
End Property

Public Sub OpenAchive()
    'Free File
    dFilePtr = FreeFile
    
    'Open the file for reading.
    Open Filename For Binary As #dFilePtr
        'Get File Header Info
        Get #dFilePtr, , m_ArcHeader
        'Tells us the file is open.
        IsOpen = True
End Sub

Public Sub AddFile(SrcFile As String)
    'Add a new file to the achive.
    
    If (Not IsOpen) Then
        'error code here
    Else
        Seek #dFilePtr, LOF(dFilePtr) + 1
        'Store some info on this file.
        fInfo.FileLength = FileLen(SrcFile)
        fInfo.Filename = GetFileTitle(SrcFile)
        'Put the file info.
        Put #dFilePtr, , fInfo
        'Put the file data.
        Put #dFilePtr, , fOpenFile(SrcFile)
        'INC Files count.
        m_ArcHeader.NOF = (m_ArcHeader.NOF + 1)
        'Add Total Filesize
        m_ArcHeader.FileSize = (m_ArcHeader.FileSize + fInfo.FileLength)
    End If
End Sub

Public Sub CreateAchive()
    'Free File
    dFilePtr = FreeFile
    'Fill in m_ArcHeader info.
    m_ArcHeader.ID = "Arc"
    m_ArcHeader.Version = 1
    m_ArcHeader.HeadSize = 15
    m_ArcHeader.NOF = 0
    m_ArcHeader.FileSize = 0
    
    'Open the file for writeing.
    Open Filename For Binary As #dFilePtr
        'Put the file header.
        Put #dFilePtr, , m_ArcHeader
        'Set isOpen to true
        IsOpen = True
End Sub

Public Sub DeleteFile(Index As Long)
On Error GoTo DelFlag:
Dim sTmpName As String
Dim fp As Long
Dim Offset As Long
Dim Count As Long
Dim mBytes() As Byte

    If (Not m_IsOpen) Then
        'error code
        Exit Sub
    Else
        fp = FreeFile
        'Create temp file
        sTmpName = m_Filename & ".tmp"
        'Open Temp file.
        Open sTmpName For Binary As #fp
            'Put the header
            Put #fp, , m_ArcHeader
            'File data start.
            Seek #dFilePtr, m_ArcHeader.HeadSize
            'Return Filename
            For Count = 0 To (FileCount - 1)
                'Get File info.
                Get #dFilePtr, , fInfo
                'File Current offset.
                Offset = Seek(dFilePtr) + fInfo.FileLength
                'Seek to the new offset.
                Seek #dFilePtr, Offset
                'Check if the file is selected for delete.
                If (Index <> Count) Then
                    'Resize mBytes to hold file bytes.
                    ReDim Preserve mBytes(0 To fInfo.FileLength - 1) As Byte
                    'Get the file data.
                    Get #dFilePtr, (Offset - fInfo.FileLength), mBytes
                    'Put file info header
                    Put #fp, , fInfo
                    'Put File Bytes.
                    Put #fp, , mBytes
                Else
                    m_ArcHeader.FileSize = (m_ArcHeader.FileSize - fInfo.FileLength)
                End If
            Next Count
            'DEC count.
            m_ArcHeader.NOF = (m_ArcHeader.NOF - 1)
            'Update Header.
            Put #fp, 5, m_ArcHeader.NOF
            Put #fp, 11, m_ArcHeader.FileSize
        Close #fp
        'Close the original
        Call CloseAchive
        'Set file attr to normal
        SetAttr Filename, vbNormal
        'Delete the original File.
        Kill Filename
        'Replace the old file with the new one.
        FileCopy sTmpName, Filename
        'Reload the new file
        Call OpenAchive
        'Kill the temp file.
        Kill sTmpName
    End If
    
    'Clear up
    Erase mBytes
    Offset = 0
    Count = 0
    sTmpName = vbNullString
    
    Exit Sub
DelFlag:
    If Err Then
        MsgBox "Error:DeleteFile()" & vbCrLf & Err.Description, vbExclamation, "Error"
    End If
End Sub

Public Sub ExtractFile(Index As Long, Destination As String)
Dim Count As Long
Dim Offset As Long
Dim mBytes() As Byte

    If (Not IsOpen) Then
        'error code
        Exit Sub
    Else
        'File data start.
        Seek #dFilePtr, m_ArcHeader.HeadSize
        'Get the Fileinfo we need to extract.
        For Count = 0 To Index
            Get #dFilePtr, , fInfo
            'File Current offset.
            Offset = Seek(dFilePtr) + fInfo.FileLength
            'Seek to the new offset.
            Seek #dFilePtr, Offset
        Next Count
    End If

    If (fInfo.FileLength = 0) Then
        Exit Sub
    Else
        'Resize mBytes to hold file bytes.
        ReDim Preserve mBytes(0 To fInfo.FileLength - 1) As Byte
        'Get the file data.
        Get #dFilePtr, (Offset - fInfo.FileLength), mBytes
        'Write bytes back to new file.
        Call fPutBytes(Destination, mBytes)
        
        'Clean vars.
        Offset = 0
        Count = 0
        Erase mBytes
    End If
End Sub

Public Sub ExtractAll(Destination As String)
Dim Count As Long
Dim Offset As Long
Dim mBytes() As Byte
On Error Resume Next

    If (Not IsOpen) Then
        'error code here
        Exit Sub
    Else
        'File data start.
        Seek #dFilePtr, m_ArcHeader.HeadSize
        For Count = 0 To (FileCount - 1)
            'Get File info.
            Get #dFilePtr, , fInfo
            'File Current offset.
            Offset = Seek(dFilePtr) + fInfo.FileLength
            'Resize mBytes to hold file bytes.
            ReDim Preserve mBytes(0 To fInfo.FileLength - 1) As Byte
            'Get the File Bytes.
            Get #dFilePtr, , mBytes
            'Seek to the new offset.
            Seek #dFilePtr, Offset
            'Write bytes back to new file.
            fPutBytes Destination & fInfo.Filename, mBytes
        Next Count
    End If
    
    'Clear up
    Erase mBytes
    Count = 0
    Offset = 0
    
End Sub

Public Function GetFileInfo(Index As Long) As cFileInfo
Dim Idx As Long
Dim Tmp As cFileInfo
Dim Offset As Long

    If (Not IsOpen) Then
        'error code
        Exit Function
    Else
        'File data start.
        Seek #dFilePtr, m_ArcHeader.HeadSize
        'Return Filename
        For Idx = 0 To Index
            'Get File info.
            Get #dFilePtr, , fInfo
            'File Current offset.
            Offset = Seek(dFilePtr) + fInfo.FileLength
            'Seek to the new offset.
            Seek #dFilePtr, Offset
        Next Idx
    End If
    
    Set Tmp = New cFileInfo
    'Set tmp object with fInfo
    Tmp.FileLength = fInfo.FileLength
    Tmp.Filename = fInfo.Filename
    'Send back the object.
    Set GetFileInfo = Tmp
    
    'Clean vars
    Set Tmp = Nothing
    Idx = 0
End Function

Public Sub CloseAchive()
    'Close the open achive.
    If (IsOpen) Then
        'Close the file
        Close #dFilePtr
        Call Reset
        IsOpen = False
    End If
End Sub

Private Sub Reset()
    With m_ArcHeader
        .FileSize = 0
        .HeadSize = 0
        .ID = vbNullString
        .NOF = 0
        .Version = 0
    End With
End Sub

Public Sub UpDateHeader()
    'Update the file header.
    Put #dFilePtr, 5, m_ArcHeader.NOF
    'Put Filesize
    Put #dFilePtr, 11, m_ArcHeader.FileSize
End Sub

Private Function GetFileTitle(lpSrc As String) As String
Dim s_pos As Integer
    s_pos = InStrRev(lpSrc, "\", Len(lpSrc), vbBinaryCompare)
    
    If (s_pos) Then
        GetFileTitle = Mid(lpSrc, s_pos + 1)
    Else
        GetFileTitle = lpSrc
    End If
End Function

Private Function fOpenFile(lzFile As String) As Byte()
Dim fp As Integer
Dim mBytes() As Byte

    'Free File.
    fp = FreeFile
    'Open File for reading.
    Open lzFile For Binary As #fp
        If LOF(fp) <> 0 Then
            'Resize the array to hold file contents.
            ReDim Preserve mBytes(0 To LOF(fp) - 1)
        End If
        'Get File Bytes
        Get #fp, , mBytes
        'Close file.
    Close #fp
    
    'Send back the data.
    fOpenFile = mBytes
    Erase mBytes
    
End Function

Private Sub fPutBytes(SrcFile As String, ByRef Data() As Byte)
Dim fp As Integer
    'Free file.
    fp = FreeFile
    'Open the file for writeing.
    Open SrcFile For Binary As #fp
        'Put contents into the new file.
        Put #fp, , Data
    'Close file.
    Close #fp
End Sub

Private Sub Class_Terminate()
    Call Reset
    Filename = vbNullString
End Sub
