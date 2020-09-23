VERSION 5.00
Begin VB.Form frmSfx 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple File Achiver Extractor"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "frmSfx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   5355
      TabIndex        =   4
      Top             =   1395
      Width           =   5355
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   365
         Left            =   4215
         TabIndex        =   7
         Top             =   105
         Width           =   1035
      End
      Begin VB.CommandButton cmdabout 
         Caption         =   "&About"
         Height          =   365
         Left            =   2940
         TabIndex        =   6
         Top             =   105
         Width           =   1035
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Height          =   365
         Left            =   1815
         TabIndex        =   5
         Top             =   105
         Width           =   1035
      End
      Begin VB.Line lnSpacer 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   765
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line lnSpacer 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   0
         X2              =   765
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.CommandButton cmdOpen1 
      Caption         =   " ...."
      Height          =   375
      Left            =   4635
      TabIndex        =   2
      Top             =   435
      Width           =   510
   End
   Begin VB.TextBox TxtExt 
      Height          =   350
      Left            =   240
      TabIndex        =   1
      Top             =   450
      Width           =   4320
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   45
   End
   Begin VB.Label lblExt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extract To:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   780
   End
End
Attribute VB_Name = "frmSfx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private m_TmpFile As String
Private myArc As New dFileArc

Private Sub cmdabout_Click()
    MsgBox frmSfx.Caption & vbCrLf & vbCrLf & "Built with DM Simple FileAchiver.", vbInformation, "About"
End Sub

Private Sub cmdExit_Click()
    Unload frmSfx
End Sub

Private Sub cmdOpen1_Click()
    m_ExtractDir = FixPath(GetFolder(frmSfx.hwnd, "Extract To:"))
    
    If Len(m_ExtractDir) <> 1 Then
        TxtExt.Text = m_ExtractDir
    End If
    
End Sub

Private Sub cndExit_Click()

End Sub

Private Sub cmdStart_Click()
Dim cInfo As cFileInfo
Dim fCounter As Long
On Error Resume Next

    If Not FindFile(m_TmpFile) Then
        MsgBox "Data File Not Found", vbExclamation, "File Not Found"
        Unload frmSfx
        Exit Sub
    Else
        'Check if extract folder if found.
        If Not (Len(Dir(m_ExtractDir, vbDirectory)) <> 0) Then
            'Create extract folder.
            Call CreateDir(m_ExtractDir)
        End If
        
        'Extract the achive file.
        With myArc
            .Filename = m_TmpFile
            .OpenAchive
            For fCounter = 0 To .FileCount - 1
                Set cInfo = .GetFileInfo(fCounter)
                .ExtractFile fCounter, (m_ExtractDir & cInfo.Filename)
                'Add a small delay
                Sleep 10
                lblStatus.Caption = "Extracting: " & cInfo.Filename
            Next fCounter
            'Close the achive
            .CloseAchive
        End With
    End If

    'Delete achive temp file.
    Kill m_TmpFile
    Set myArc = Nothing
    Set cInfo = Nothing
    fCounter = 0
    'Clear up
    m_ExtractDir = vbNullString
    
    Call MsgBox("All files have been extracted.", vbInformation, "Finished")
    
    'Extract back to the system.
    Call cmdExit_Click
End Sub

Private Sub Form_Load()
Dim MainExe As String
Dim fp As Long
Dim Bytes() As Byte
Dim Sig As String * 3
Dim fStart As Long
On Error GoTo ReadErr:

    fp = FreeFile
    MainExe = FixPath(App.Path) & App.EXEName & ".exe"
    m_ExtractDir = FixPath(App.Path) & GetFilePart(MainExe, fFileTitle) & "\"
    'Temp Achive extract path.
    m_TmpFile = FixPath(GetTempDir) & GetFilePart(MainExe, fFileTitle) & ".sfa"
    
    Open MainExe For Binary As #fp
        Get #fp, LOF(fp) - 3, fStart
        If (fStart = 0) Then
            'error
            Err.Raise 58
        Else
            Get #fp, fStart, Sig
            If (Sig <> "Arc") Then
                Err.Raise 58
            Else
                'Resize and extract the data
                Seek #fp, fStart
                ReDim Preserve Bytes((LOF(fp) - fStart - 4))
                Get #fp, , Bytes
                'Extract the New Data.
                Open m_TmpFile For Binary As #2
                    Put #2, , Bytes
                Close #2
            End If
        End If
    Close #fp
    'Set extract textbox text
    TxtExt.Text = m_ExtractDir
    'Clear up
    MainExe = vbNullString
    Erase Bytes
    
    Exit Sub
    'Error flag
ReadErr:
    MsgBox "File Read Error.", vbExclamation, "Read Error_58"
    MainExe = vbNullString
    Close #fp
    Unload frmSfx
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSfx = Nothing
End Sub

Private Sub pBottom_Resize()
    lnSpacer(0).X2 = pBottom.ScaleWidth
    lnSpacer(1).X2 = pBottom.ScaleWidth
End Sub
