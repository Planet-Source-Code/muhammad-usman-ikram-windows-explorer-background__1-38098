VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFolder 
      Caption         =   "Select Folder Where You Want to See Image"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   4695
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdBrows 
      Caption         =   "Browse"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   480
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Comments:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email: usman@ematic.com"
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   1560
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Muhammad Usman Ikram"
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Brows Button to Select GIF or BMP image file name"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4020
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strResFolder As String ''Variable to sho image in folder background
Dim lpFileName As String ''For IniFile name
Dim ret As Long
Private Sub cmdApply_Click()

lpFileName = strResFolder & "Desktop.ini"
    
ret = WritePrivateProfileString("ExtShellFolderViews", "{BE098140-A513-11D0-A3A4-00C04FD706EC}", "{BE098140-A513-11D0-A3A4-00C04FD706EC}", lpFileName)
ret = WritePrivateProfileString("{BE098140-A513-11D0-A3A4-00C04FD706EC}", "Attributes", "1", lpFileName)
ret = WritePrivateProfileString("{BE098140-A513-11D0-A3A4-00C04FD706EC}", "IconArea_Image", cDialog.FileName, lpFileName)
ret = WritePrivateProfileString(".ShellClassInfo", "ConfirmFileOp", "0", lpFileName)

ret = WritePrivateProfileString("Filepath", "Filename", lpFileName, lpFileName)
MsgBox "Process Complete Successfully" & vbCrLf & "Now goto " & strResFolder & " and see Image. If there is no image then press F5 key", vbInformation, ""

End Sub

Private Sub cmdBrows_Click()
    With cDialog
        .CancelError = True
        .Flags = cdlOFNHideReadOnly
        .Filter = "Image File (*.bmp, *.gif)|*.bmp;*.gif"
        .DialogTitle = "Image File"
        .InitDir = "C:\"
        On Error GoTo er
        .ShowOpen
    End With
    lblPath.Caption = cDialog.FileName
If cDialog.FileName = "" Then
Else
    cmdFolder.Enabled = True
End If

er:
    Resume Next
End Sub

Private Sub cmdFolder_Click()
strResFolder = BrowseForFolder(hWnd, "Please select a folder.")

If strResFolder = "" Then
'    Call MsgBox("The Cancel button was pressed.", vbExclamation)
Else
    cmdApply.Enabled = True
'    Call MsgBox("The folder " & strResFolder & " was selected.", vbExclamation)
End If

End Sub
