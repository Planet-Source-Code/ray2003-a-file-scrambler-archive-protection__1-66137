VERSION 5.00
Begin VB.Form frmDescrDest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extract destination folder:"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   4800
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   5775
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   5775
   End
   Begin VB.CommandButton cmdCreateDir 
      Caption         =   "Create new dir"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh listing"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblLocation 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   5775
   End
End
Attribute VB_Name = "frmDescrDest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdCreateDir_Click()
    Dim s As String
    s = InputBox("Name of new directory:")
    If s <> "" Then
        If InStr(s, "\") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, "/") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, ":") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, "*") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, "?") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, """") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, "<") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, ">") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, "|") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        MkDir AddSep(Dir1.Path) & s
        Dir1.Refresh
    End If
End Sub

Private Sub cmdOK_Click()
    If lblLocation.Caption <> "" Then
        ListDirX lblLocation.Caption
        If lFilecount > 0 Or lDircount > 0 Then
            If MsgBox("Selected directory is NOT empty, do you still want to proceed?", vbYesNo Or vbExclamation) = vbYes Then
                frmMain.lblDestination.Caption = lblLocation.Caption
                Me.Hide
            End If
        Else
            frmMain.lblDestination.Caption = lblLocation.Caption
            Me.Hide
        End If
    Else
        MsgBox "No directory selected", vbExclamation
    End If
End Sub

Private Sub cmdRefresh_Click()
    Dir1.Refresh
End Sub

Private Sub Dir1_Change()
    lblLocation.Caption = AddSep(Dir1.Path)
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Dir1.Refresh
End Sub

