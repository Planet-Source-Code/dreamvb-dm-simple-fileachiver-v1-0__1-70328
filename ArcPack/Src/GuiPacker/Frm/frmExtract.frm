VERSION 5.00
Begin VB.Form frmExt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extract"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   3795
      TabIndex        =   3
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "&Extract"
      Height          =   375
      Left            =   2445
      TabIndex        =   2
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen1 
      Caption         =   "...."
      Height          =   390
      Left            =   4515
      TabIndex        =   1
      Top             =   420
      Width           =   510
   End
   Begin VB.TextBox txtExtract 
      Height          =   360
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   435
      Width           =   4260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   165
      X2              =   5025
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   165
      X2              =   5025
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblExtract 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extract To:"
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   195
      Width           =   780
   End
End
Attribute VB_Name = "frmExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload frmExt
End Sub

Private Sub cmdExtract_Click()
    m_CancelButton = False
    Unload frmExt
End Sub

Private Sub cmdOpen1_Click()
    'Extract Path.
    m_ExtractDir = FixPath(GetFolder(frmExt.hWnd, "Extract To:"))
    
    If (m_ExtractDir <> "\") Then
        txtExtract.Text = m_ExtractDir
    End If
End Sub

Private Sub Form_Load()
    Set frmExt.Icon = Nothing
    txtExtract.Text = m_ExtractDir
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmExt = Nothing
End Sub
