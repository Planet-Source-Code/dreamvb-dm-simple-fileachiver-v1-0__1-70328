Attribute VB_Name = "dTools"
Option Explicit

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Enum TFileParts
    fDrive = 0
    fPath = 1
    fFileName = 2
    fFileTitle = 3
    fFileExt = 4
    fFullPathNoFileExt = 5
End Enum

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40

Public m_CancelButton As Boolean
Public m_ExtractDir As String

Public Function OpenFile(lzFile As String) As Byte()
Dim fp As Long
Dim fBytes() As Byte
    fp = FreeFile
    
    Open lzFile For Binary As #fp
        If LOF(fp) <> 0 Then
            ReDim Preserve fBytes(0 To LOF(fp) - 1)
        End If
        Get #fp, , fBytes
    Close #fp
    
    OpenFile = fBytes
    Erase fBytes
End Function

Public Function FindFile(lzFileName As String) As Boolean
    If Trim(Len(lzFileName)) = 0 Then Exit Function
    FindFile = LenB(Dir(lzFileName)) <> 0
End Function

Public Function FixPath(lzPath As String) As String
    If Right$(lzPath, 1) = "\" Then
        FixPath = lzPath
    Else
        FixPath = lzPath & "\"
    End If
End Function

Public Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String)
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim Offset As Integer
    
    With bInf
        .hOwner = hWndOwner
        .lpszTitle = sTitle
        .ulFlags = (BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE)
        
        PathID = SHBrowseForFolder(bInf)
        RetPath = Space(512)
        RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
        
        If (RetVal) Then
            Offset = InStr(RetPath, Chr$(0))
            GetFolder = Left$(RetPath, Offset - 1)
        End If
    End With
    
End Function

Public Function GetFilePart(lzFile As String, FilePart As TFileParts) As String
On Error Resume Next
Dim xPos As Integer
Dim yPos As Integer
Dim zPos As Integer
Dim sTmp As String
    
    'Returns Parts of a filename or path

    'Get delimiter positions
    xPos = InStr(1, lzFile, ":\", vbBinaryCompare)
    yPos = InStrRev(lzFile, "\", Len(lzFile), vbBinaryCompare)
    zPos = InStrRev(lzFile, ".", Len(lzFile), vbBinaryCompare)
        
    Select Case FilePart
        Case fDrive
            'Get Drive
            If (xPos) Then
                sTmp = Left$(lzFile, 3)
            End If
        Case fFileName
            'Get Filename
            If (yPos) Then
                sTmp = Mid$(lzFile, yPos + 1)
            End If
        Case fFileExt
            'Get File Ext
            If (zPos) Then
                sTmp = LCase$(Mid(lzFile, zPos))
            End If
        Case fFileTitle
            'Get File Title
            If (yPos) And (zPos) Then
                sTmp = Mid$(lzFile, (yPos + 1), (zPos - yPos) - 1)
            End If
        Case fPath
            'Get Path
            If (yPos) Then
                sTmp = Left$(lzFile, yPos)
            End If
        Case fFullPathNoFileExt
            'Returns full Drive, Path File Title with no File ext
            'eg C:\windows\notepad.exe, returns C:\windows\notepad
            If (zPos) Then
                sTmp = Left$(lzFile, zPos - 1)
            End If
    End Select
    
    GetFilePart = sTmp
    
    sTmp = vbNullString
    
End Function

Public Sub CreateDir(PathName As String)
Dim vLst() As String
Dim Counter As Long
Dim Path As String
Dim FolderPath As String
Dim sTmp As String

    'Append Backslash if needed
    sTmp = FixPath(PathName)
    
    'Split paths
    vLst = Split(sTmp, "\", , vbBinaryCompare)

    Do While Counter < UBound(vLst)
        'Append Backslash if needed
        Path = FixPath(vLst(Counter))
        
        If Len(Path) <> 0 Then
            'Build the folder path
            FolderPath = (FolderPath & Path)
            If LenB(Dir(FolderPath, vbDirectory)) = 0 Then
                'Create FolderPath if not found.
                MkDir FolderPath
            End If
        End If
        'INC Counter , get next path.
        Counter = (Counter + 1)
    Loop
    
    Erase vLst
    Counter = 0
    Path = vbNullString
    FolderPath = vbNullString
    sTmp = vbNullString
    
End Sub

Public Function GetTempDir() As String
Dim iRet As Long
Dim sBuff As String
    sBuff = Space(255)
    iRet = GetTempPath(255, sBuff)
    If (iRet) Then
        GetTempDir = Left$(sBuff, iRet - 1)
    End If
    
    sBuff = vbNullString
End Function

Public Sub ExecFile(dHwnd As Long, lzFileName As String)
     ShellExecute dHwnd, "open", lzFileName, vbNullString, vbNullString, 1
End Sub
