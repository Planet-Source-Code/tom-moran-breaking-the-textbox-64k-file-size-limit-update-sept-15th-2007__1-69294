VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHugeText 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Breaking the TextBox 64K Limit"
   ClientHeight    =   5220
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7560
   Icon            =   "frmHugeText.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   7260
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton optLoadFile 
      Caption         =   "API Method (Unlimited file size)"
      Height          =   315
      Index           =   1
      Left            =   4020
      TabIndex        =   2
      Top             =   60
      Value           =   -1  'True
      Width           =   2715
   End
   Begin VB.OptionButton optLoadFile 
      Caption         =   "Normal VB Method (64K limit)"
      Height          =   315
      Index           =   0
      Left            =   540
      TabIndex        =   1
      Top             =   60
      Width           =   2595
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHugeText.frx":058A
      Top             =   480
      Width           =   7215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open File..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save File..."
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmHugeText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuExit_Click()
 Unload Me
End Sub

Private Sub mnuOpen_Click()
  FileOpenProc
End Sub

Private Sub mnuSave_Click()
   Dim strSaveFileName As String
        strSaveFileName = GetFileName("Untitled.txt")
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
End Sub

