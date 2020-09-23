VERSION 5.00
Begin VB.Form frmLocateArchive 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Locate scrambled archive:"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox fil 
      Height          =   480
      Left            =   120
      Pattern         =   "0000.fs-data"
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   5775
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5775
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblOK 
      Alignment       =   2  'Center
      Caption         =   "Scrambled archive file present"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   5775
   End
End
Attribute VB_Name = "frmLocateArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If lblOK.Visible Then
        frmMain.lblScrambledFile.Caption = AddSep(fil.Path) & "0000.fs-data"
        Me.Hide
    Else
        MsgBox "No scrambled archive files present in the current directory", vbExclamation
    End If
End Sub

Private Sub Dir1_Change()
    fil.Path = Dir1.Path
    If fil.ListCount = 0 Then
        lblOK.Visible = False
    Else
        lblOK.Visible = True
    End If
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Dir1.Refresh
End Sub

