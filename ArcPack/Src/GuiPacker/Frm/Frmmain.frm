VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "DM Simple FileAchiver"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7350
   Icon            =   "Frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   1245
      Top             =   2115
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmain.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstV 
      Height          =   750
      Left            =   0
      TabIndex        =   5
      Top             =   975
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImgLst"
      SmallIcons      =   "ImgLst"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FileType"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FileSize"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox pBar2 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   490
      TabIndex        =   3
      Top             =   630
      Width           =   7350
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   4
         Top             =   75
         Width           =   570
      End
   End
   Begin VB.PictureBox pBar1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   0
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   490
      TabIndex        =   1
      Top             =   0
      Width           =   7350
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImgToolBar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "NEW"
               Object.ToolTipText     =   "Create New Achive"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "OPEN"
               Object.ToolTipText     =   "Open Achive"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "ADD"
               Object.ToolTipText     =   "Add Files to Achive"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "EXTRACT"
               Object.ToolTipText     =   "Extract Files from Achive"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "DELETE"
               Object.ToolTipText     =   "Delete Files from Achive"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar sBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   2835
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   2381
            MinWidth        =   2381
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12462
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgToolBar 
      Left            =   645
      Top             =   2115
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmain.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmain.frx":12AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmain.frx":1F00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmain.frx":2B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmain.frx":37A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   150
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   2048
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Archive"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Archive"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close Archive"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBlank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuHidden 
      Caption         =   "#"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen2 
         Caption         =   "Open"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExt2 
         Caption         =   "Extract"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDel2 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select &All"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "Action"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add File"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete File"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "&Extract"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBlank0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll2 
         Caption         =   "Select All"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuild 
         Caption         =   "Build SFX"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawEdge Lib "user32.dll" (ByVal hdc As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private sfxFile As String
Private MyArc As New dFileArc

Private Sub ClickButton(Index As Integer)
Dim mnuButton As Button
    Set mnuButton = Toolbar1.Buttons(Index)
    Call Toolbar1_ButtonClick(mnuButton)
    Set mnuButton = Nothing
End Sub

Private Sub UpdateStatus()
    lblFileName.Caption = MyArc.Filename
    sBar1.Panels(1).Text = "Files: " & MyArc.FileCount
    sBar1.Panels(2).Text = "Size: " & MyArc.FileLength
    sBar1.Panels(1).Visible = True
    sBar1.Panels(2).Visible = True
    'Enable / Disable Add, Extract Toolbar buttons.
    Toolbar1.Buttons(3).Enabled = MyArc.IsOpen 'Add
    Toolbar1.Buttons(4).Enabled = MyArc.FileCount 'Extract
    Toolbar1.Buttons(5).Enabled = MyArc.FileCount 'Delete
    'MenuItems
    mnuClose.Enabled = MyArc.IsOpen 'Close
    mnuAdd.Enabled = MyArc.IsOpen 'Add
    mnuDelete.Enabled = MyArc.FileCount 'Delete
    mnuExtract.Enabled = MyArc.FileCount 'Extract
    
    mnuOpen2.Enabled = mnuDelete.Enabled
    mnuDel2.Enabled = mnuDelete.Enabled
    mnuExt2.Enabled = mnuExtract.Enabled
    mnuSelAll2.Enabled = MyArc.FileCount
    mnuSelAll.Enabled = mnuSelAll2.Enabled
    mnuBuild.Enabled = mnuSelAll2.Enabled
    
    'Resize the Column headers
    Call LVResizeColumn(LstV, 0)
    Call LVResizeColumn(LstV, 1)
    Call LVResizeColumn(LstV, 2)
End Sub

Private Sub ExtractFiles()
Dim lItem As ListItem
Dim fInfo As cFileInfo
Dim lIndex As Long
Dim FullFile As String
    
    If (LstV.ListItems.Count) Then
        'Create folders if needed
        Call CreateDir(m_ExtractDir)
    End If
    
    For Each lItem In LstV.ListItems
        'See witch item is selected.
        If (lItem.Selected) Then
            'Get the File Index.
            lIndex = (lItem.Index - 1)
            'Get the file info.
            Set fInfo = MyArc.GetFileInfo(lIndex)
            'Build full Path/File extract path
            FullFile = m_ExtractDir & fInfo.Filename
            'Extract the Files
            MyArc.ExtractFile lIndex, FullFile
        End If
    Next lItem
    
    Set fInfo = Nothing
    Set lItem = Nothing
    FullFile = vbNullString
    lIndex = 0
    
End Sub

Function NewDLG() As Boolean
On Error GoTo NewErr:
    With CD1
        .CancelError = True
        .DialogTitle = "New Achive"
        .Filter = "Simple File Achiver Files(*.sfa)|*.sfa|"
        .Filename = vbNullString
        .ShowSave
        'Close the current achive if already open.
        If (MyArc.IsOpen) Then
            MyArc.CloseAchive
        End If
        'Clear Listview.
        LstV.ListItems.Clear
        'Create new Achive.
        MyArc.Filename = .Filename
        MyArc.CreateAchive
        'Update statusbar.
        Call UpdateStatus
        'Default Extract path
        m_ExtractDir = FixPath(GetFilePart(.Filename, fFullPathNoFileExt))
    End With

Exit Function
'Error flag
NewErr:
    If Err Then
        'Cancel was pressed.
        NewDLG = True
    End If
End Function

Private Sub AddFilesDLG()
On Error GoTo AddErr:
Dim vFiles() As String
Dim Count As Long
Dim nSize As Long
Dim lzAddPath As String
Dim lzAddFile As String
Dim fInfo As New cFileInfo

    With CD1
        .CancelError = True
        .DialogTitle = "Add Files"
        .Filter = "All Files(*.*)|*.*|"
        .Flags = (cdlOFNAllowMultiselect Or cdlOFNExplorer)
        .Filename = vbNullString
        .ShowOpen
        
        vFiles = Split(.Filename, Chr(0), , vbBinaryCompare)
        nSize = UBound(vFiles)
        
        If (nSize = 0) Then
            'Add Single File.
            MyArc.AddFile .Filename
            fInfo.Filename = GetFilePart(.Filename, fFileName)
            fInfo.FileLength = FileLen(.Filename)
            Call LVAddFileInfo(LstV, fInfo.Filename, fInfo.FileLength)
        Else
            'Get the path.
            lzAddPath = FixPath(vFiles(0))
            For Count = 1 To nSize
                'Get the filename.
                lzAddFile = (lzAddPath & vFiles(Count))
                'Add file to achive.
                MyArc.AddFile lzAddFile
                'Get the File Info
                fInfo.Filename = GetFilePart(lzAddFile, fFileName)
                fInfo.FileLength = FileLen(lzAddFile)
                Call LVAddFileInfo(LstV, fInfo.Filename, fInfo.FileLength)
            Next Count
        End If
    End With
    
    'Update File Headers
    MyArc.UpDateHeader
    'Update status display
    LstV.ListItems(1).Selected = False
    Call UpdateStatus
    
    Set fInfo = Nothing
    
    Exit Sub
    'Error Flag
AddErr:
    If Err Then
        Err.Clear
    End If
End Sub

Private Sub OpenDLG()
Dim Count As Long
On Error GoTo OpenErr:
Dim fInfo As cFileInfo

    With CD1
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = "Simple File Achiver Files(*.sfa)|*.sfa|"
        .Flags = 0
        .Filename = vbNullString
        .ShowOpen
        'Close achive of it already open.
        If (MyArc.IsOpen) Then
            MyArc.CloseAchive
        End If
        'Assign the new filename to open.
        MyArc.Filename = .Filename
        'Open the achive.
        MyArc.OpenAchive
        'Clear listview items.
        LstV.ListItems.Clear
        'Check achive header
        If (MyArc.Signature <> "Arc") Then
            MyArc.CloseAchive
            lblFileName.Caption = vbNullString
            sBar1.Panels(1).Visible = False
            sBar1.Panels(2).Visible = False
        Else
            'Default Extract path
            m_ExtractDir = FixPath(GetFilePart(.Filename, fFullPathNoFileExt))
            'Add the files to the listview control.
            For Count = 0 To MyArc.FileCount - 1
               Set fInfo = MyArc.GetFileInfo(Count)
               Call LVAddFileInfo(LstV, fInfo.Filename, fInfo.FileLength)
            Next Count
        End If
    End With
    
    'Update statusbar
    Call UpdateStatus
    LstV.ListItems(1).Selected = False
    Set fInfo = Nothing
    
    Exit Sub
OpenErr:
    If Err Then
        Err.Clear
    End If
End Sub

Private Sub DrawRect(PicBox As PictureBox, RectType As Integer)
Dim rc1 As RECT
    With PicBox
        .Cls
        SetRect rc1, 0, 0, .ScaleWidth, .ScaleHeight
        DrawEdge .hdc, rc1, RectType, &HF
        SetRect rc1, 1, 1, .ScaleWidth - 1, .ScaleHeight - 1
        DrawEdge .hdc, rc1, (2 * RectType), &HF
    End With
End Sub

Private Sub Form_Load()
Dim vLst() As String
    sfxFile = FixPath(App.Path) & "SelfExt.exe"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Close Archive if open.
    If (MyArc.IsOpen) Then
        MyArc.CloseAchive
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    LstV.Width = (frmmain.ScaleWidth - LstV.Left) - Screen.TwipsPerPixelX
    LstV.Height = (frmmain.ScaleHeight - sBar1.Height - LstV.Top) - Screen.TwipsPerPixelY
    'Resize List column
    Call LVResizeColumn(LstV, 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmExt = Nothing
    Set frmabout = Nothing
    Set frmmain = Nothing
End Sub

Private Sub LstV_DblClick()
Dim cInfo As cFileInfo
Dim Idx As Long
Dim TmpFile As String

    If (LstV.ListItems.Count = 0) Then
        Exit Sub
    Else
        Idx = (LstV.SelectedItem.Index - 1)
        'Extract selected filename.
        Set cInfo = MyArc.GetFileInfo(Idx)
        'Create temp extract filename
        TmpFile = FixPath(GetTempDir) & cInfo.Filename
        If FindFile(TmpFile) Then
            SetAttr TmpFile, vbNormal
            Kill TmpFile
        End If
        'Extract temp file
        Call MyArc.ExtractFile(Idx, TmpFile)
        'Execute the file.
        Call ExecFile(frmmain.hwnd, TmpFile)
        'Clear up.
        Idx = 0
        TmpFile = vbNullString
        Set cInfo = Nothing
    End If
    
End Sub

Private Sub LstV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button <> vbLeftButton) Then
        PopupMenu MnuHidden
    End If
End Sub

Private Sub mnuAbout_Click()
    frmabout.Show vbModal, frmmain
End Sub

Private Sub mnuAdd_Click()
    Call ClickButton(3)
End Sub

Private Sub mnuBuild_Click()
Dim Ans As Integer
Dim fp As Long
Dim Offset As Long
Dim TmpFile As String

    If Not FindFile(sfxFile) Then
        MsgBox "Sfx File Not Found.", vbExclamation, "File Not Found"
        Exit Sub
    End If
    
    'Sfx filename to create.
    TmpFile = GetFilePart(MyArc.Filename, fFullPathNoFileExt) & ".exe"
    'Create a copy of the src sfx file.
    Call FileCopy(sfxFile, TmpFile)
    'Get FreeFile.
    fp = FreeFile
    Open TmpFile For Binary As #fp
        'Offset
        Offset = LOF(fp)
        'Move to end of file.
        Seek #fp, Offset
        'Put the archive data.
        Put #fp, , OpenFile(MyArc.Filename)
        'Place start offset
        Put #fp, , Offset
    Close #fp
    
    Ans = MsgBox(TmpFile & " Has been built from " & vbCrLf & MyArc.Filename _
    & vbCrLf & vbCrLf & "Do you want to test the Sfx file now.", vbInformation Or vbYesNo, "Finished")
    
    'Ask user if thay want to test the file.
    If (Ans = vbYes) Then
        Call ExecFile(frmmain.hwnd, TmpFile)
    End If
    
    'Clean up
    TmpFile = vbNullString
End Sub

Private Sub mnuClose_Click()
    'Close the archive
    If (MyArc.IsOpen) Then
        LstV.ListItems.Clear
        MyArc.Filename = vbNullString
        m_ExtractDir = vbNullString
        MyArc.CloseAchive
        Call UpdateStatus
        sBar1.Panels(1).Visible = False
        sBar1.Panels(2).Visible = False
    End If
End Sub

Private Sub mnuDel2_Click()
    Call mnuDelete_Click
End Sub

Private Sub mnuDelete_Click()
    Call ClickButton(5)
End Sub

Private Sub mnuExit_Click()
    Call mnuClose_Click
    Unload frmmain
End Sub

Private Sub mnuExt2_Click()
    Call ClickButton(4)
End Sub

Private Sub mnuExtract_Click()
    Call ClickButton(4)
End Sub

Private Sub mnuNew_Click()
    Call ClickButton(1)
End Sub

Private Sub mnuOpen_Click()
    Call ClickButton(2)
End Sub

Private Sub mnuOpen2_Click()
    Call LstV_DblClick
End Sub

Private Sub mnuSelAll_Click()
    Call LVSelectSelectAll(LstV)
End Sub

Private Sub mnuSelAll2_Click()
    Call mnuSelAll_Click
End Sub

Private Sub pBar1_Resize()
    DrawRect pBar1, 2
End Sub

Private Sub pBar2_Resize()
    DrawRect pBar2, 2
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Ans As Integer
Dim lCount As Long

    Select Case Button.Key
        Case "NEW"
            If (Not NewDLG) Then
                'Show add files dialog.
                Call AddFilesDLG
            End If
        Case "OPEN"
            Call OpenDLG
        Case "ADD"
            Call AddFilesDLG
        Case "EXTRACT"
            'Set default cancel to true.
            m_CancelButton = True
            'Show extract dialog.
            frmExt.Show vbModal, frmmain
            
            If (Not m_CancelButton) Then
                'Extract Selected Files
                If (LVHasSelectedItems(LstV)) Then
                    Call ExtractFiles
                Else
                    'Create folders if needed
                    Call CreateDir(m_ExtractDir)
                    'Extract all files to folder above.
                    Call MyArc.ExtractAll(m_ExtractDir)
                End If
            End If
        Case "DELETE"
            If (Not LVHasSelectedItems(LstV)) Then
                MsgBox "Nothing selected.", vbExclamation, "Nothing to Delete"
                Exit Sub
            Else
                'Ask user if they want to delete the item.
                Ans = MsgBox("Are you sure you want to delete the selected files.", vbYesNo Or vbQuestion, "Delete")
            End If
            
            If (Ans = vbYes) Then
                With LstV
                    For lCount = .ListItems.Count To 1 Step -1
                        If (.ListItems(lCount).Selected) Then
                            .ListItems.Remove lCount
                            Call MyArc.DeleteFile(lCount - 1)
                        End If
                    Next lCount
                End With
                'Update status
                Call UpdateStatus
            End If
    End Select
    
End Sub
