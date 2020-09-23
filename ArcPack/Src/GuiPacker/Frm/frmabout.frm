VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   3495
      TabIndex        =   1
      Top             =   0
      Width           =   3495
      Begin VB.Line ln1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   555
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM Simple FileAchiver"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   825
         TabIndex        =   2
         Top             =   165
         Width           =   2400
      End
      Begin VB.Image ImgIcon 
         Height          =   480
         Left            =   150
         Picture         =   "frmabout.frx":0000
         Top             =   75
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   365
      Left            =   2655
      TabIndex        =   0
      Top             =   1380
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Written by DreamVB."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1035
      Width           =   1830
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   765
      Width           =   930
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    Unload frmabout
End Sub

Private Sub Form_Load()
    Set frmabout.Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmabout = Nothing
End Sub

Private Sub pTop_Resize()
    ln1.X2 = pTop.ScaleWidth
End Sub
