VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "FolderIcon"
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2820
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "Icon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1395
      Left            =   120
      TabIndex        =   4
      Top             =   1620
      Width           =   4995
      Begin VB.TextBox Text_Icon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F8DABA&
         Height          =   285
         Left            =   1260
         TabIndex        =   6
         Top             =   240
         Width           =   3645
      End
      Begin VB.PictureBox Pic_Icon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   1260
         ScaleHeight     =   465
         ScaleWidth      =   540
         TabIndex        =   5
         Top             =   720
         Width           =   540
      End
      Begin VB.Label cmdClear 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   3840
         TabIndex        =   12
         Top             =   1035
         Width           =   1035
      End
      Begin VB.Label cmdIconBr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Browse ..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3840
         TabIndex        =   11
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preview :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Icon path :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Folder"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4995
      Begin VB.TextBox Foldertxt 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label cmdFldbr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Browse ..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3840
         TabIndex        =   10
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folder Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a folder and an icon please ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   3435
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   -300
      X2              =   7260
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   7560
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   5220
      X2              =   5220
      Y1              =   5280
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   5280
      Y2              =   0
   End
   Begin VB.Label cmdAbout 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   2400
      TabIndex        =   15
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label cmdCancel 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   1260
      TabIndex        =   14
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label cmdSave 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label cmdX 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "X"
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
      Height          =   190
      Left            =   4970
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   35
      Width           =   220
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FolderIcon v1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5250
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'FolderIcon v1.0
'---------------
'
'By Matthias "Thor" Stevens
'
'Original code by Juaxx
'(zyeagerr@yahoo.com,
'http://members.xoom.com/yeager/,
'http://yeagerin.tripod.com/)
'
'Interface design (VB) based on :
'PassGen v2.5 by Master Yoda
'(masteryoda@webone.com.au,
'http://webone.com.au/~jpettit)
'this interface system was also used in SubSeven
'
'---------------
'
'FolderIcon is FREEWARE and OPEN SOURCE
'
'---------------
'
'thor@ eego.net
'ICQ# : 15562140
'http://dive.to/vis-o-rama
'
'---------------



Option Explicit
'Move form without a border declarations
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSFolder = &H112

Public Function BrowseForFolder(selectedPath As String) As String
Dim Browse_for_folder As BROWSEINFOTYPE
Dim itemID As Long
Dim selectedPathPointer As Long
Dim tmpPath As String * 256
With Browse_for_folder
    .hOwner = Me.hWnd ' Window Handle
    .lpszTitle = "Browse for folders with directory pre-selection, Roman Blachman" ' Dialog Title
    .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr) ' Dialog callback function that preselectes the folder specified
    selectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1) ' Allocate a string
    CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1 ' Copy the path to the string
    .lParam = selectedPathPointer ' The folder to preselect
End With
itemID = SHBrowseForFolder(Browse_for_folder) ' Execute the BrowseForFolder API
If itemID Then
    If SHGetPathFromIDList(itemID, tmpPath) Then ' Get the path for the selected folder in the dialog
        BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1) ' Take only the path without the nulls
    End If
    Call CoTaskMemFree(itemID) ' Free the itemID
End If
Call LocalFree(selectedPathPointer) ' Free the string from the memory
End Function


Private Sub cmdAbout_Click()
About.Show
Me.Enabled = False
End Sub

Private Sub cmdAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAbout.BackColor = &HFFC0C0
End Sub

Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAbout.BackColor = &HFF8080
cmdX.BackColor = &HFF8080
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000
cmdClear.BackColor = &H800000


End Sub

Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancel.BackColor = &HFFC0C0
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancel.BackColor = &HFF8080
cmdX.BackColor = &HFF8080
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdClear.BackColor = &H800000


End Sub

Private Sub cmdClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClear.BackColor = &HFFC0C0
End Sub

Private Sub cmdClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClear.BackColor = &HFF8080
cmdX.BackColor = &HFF8080
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000
End Sub

Private Sub cmdIconBr_Click()
CommonDialog1.DialogTitle = "Select an Icon"
CommonDialog1.Filter = "Icons (*.ico)|*.ico"
CommonDialog1.ShowOpen

On Error Resume Next
    If CommonDialog1.FileName <> "" Then
        Text_Icon.Text = CommonDialog1.FileName
        Pic_Icon.Picture = LoadPicture(CommonDialog1.FileName)
    Else
        Exit Sub
    End If


End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdClear_Click()
    Pic_Icon.Picture = Me.Picture
    Text_Icon.Text = ""
End Sub

Private Sub cmdIconBr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdIconBr.BackColor = &HFFC0C0
End Sub

Private Sub cmdIconBr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdIconBr.BackColor = &HFF8080
cmdX.BackColor = &HFF8080
cmdFldbr.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdCancel.BackColor = &H800000
cmdClear.BackColor = &H800000
End Sub

Private Sub cmdSave_Click()
    Save_Configuration
End Sub


Public Sub savekey(Hkey As Long, strPath As String)
Dim KeyHandle&
    RegCreateKey Hkey, strPath, KeyHandle&
    RegCloseKey KeyHandle&
End Sub

Public Sub Save_Configuration()
Dim Folder As String
Folder = Foldertxt.Text
On Error Resume Next
Dim File_NUM As Integer, Line_Text As String, Text_Buff As String
File_NUM = FreeFile: Text_Buff = ""
'Exists...
 If StrConv(Dir(Folder & "\desktop.ini", vbSystem), vbUpperCase) <> "" Or _
     StrConv(Dir(Folder & "\desktop.ini", vbArchive), vbUpperCase) <> "" Or _
     StrConv(Dir(Folder & "\desktop.ini", vbHidden), vbUpperCase) <> "" Then
Dim num_f As Integer
num_f = FreeFile
     Open Folder & "\Desktop.ini" For Input As #num_f
      While Not EOF(num_f)
      Dim Pass_Icon As Boolean
      Pass_Icon = False
        'Replace the old lines...
            Line Input #num_f, Line_Text
            If Left(StrConv(Line_Text, vbUpperCase), 9) = "ICONFILE=" Then
                Text_Buff = Text_Buff & vbCrLf & "IconFile=" & Text_Icon.Text
                Pass_Icon = True
            ElseIf Left(StrConv(Line_Text, vbUpperCase), 8) = "ICOINDEX" Then
                Text_Buff = Text_Buff & vbCrLf & "IcoIndex=0"
            Else
                If Line_Text <> vbCrLf And Len(Line_Text) > 1 Then _
                    Text_Buff = Text_Buff & vbCrLf & Line_Text
            End If
        Wend
        Close num_f
        If Not Pass_Icon Then _
          Text_Buff = "[.ShellClassInfo]" & vbCrLf & _
            "IconFile=" & Text_Icon.Text & vbCrLf & _
            "IcoIndex=0"
    Else
        Text_Buff = "[.ShellClassInfo]" & vbCrLf & _
        "IconFile=" & Text_Icon.Text & vbCrLf & _
        "IcoIndex=0"
    End If
Write_File Folder & "\Desktop.ini", Text_Buff
'Now we've to attrib the folder +s -> system :]
' without doing this the icon won't appear !!
Attribs Folder, "+s"
End Sub

Private Sub cmdFldbr_Click()
Dim tmpPath As String
tmpPath = Foldertxt ' Take the selected path from txtStart
If Len(tmpPath) > 0 Then
    If Not Right$(tmpPath, 1) <> "\" Then tmpPath = Left$(tmpPath, Len(tmpPath) - 1) ' Remove "\" if the user added
End If
Foldertxt = tmpPath
tmpPath = BrowseForFolder(tmpPath) ' Browse for folder
If tmpPath = "" Then
    Exit Sub
Else
    Foldertxt = tmpPath ' If the user selected a folder
End If
End Sub

Private Sub cmdFldbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdFldbr.BackColor = &HFFC0C0
    
End Sub

Private Sub cmdFldbr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdFldbr.BackColor = &HFF8080
cmdX.BackColor = &HFF8080
cmdSave.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdCancel.BackColor = &H800000
cmdClear.BackColor = &H800000
End Sub


Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSave.BackColor = &HFFC0C0
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSave.BackColor = &HFF8080
cmdX.BackColor = &HFF8080
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdCancel.BackColor = &H800000
cmdClear.BackColor = &H800000
End Sub

Private Sub cmdX_Click()
End
End Sub

Private Sub cmdX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdX.BackColor = &HC00000
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000

End Sub

Private Sub Foldertxt_Change()
cmdX.BackColor = &HFF8080
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdX.BackColor = &HFF8080
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Kill "C:\r001.bat"
End Sub
Private Sub Attribs(MyFileOrDir As String, MyOptions As String)
Dim File_NUM As Integer
File_NUM = FreeFile
Open "C:\r001.bat" For Output As #File_NUM
    Print #File_NUM, "@echo off" & vbCrLf & _
     "attrib " & MyOptions & " """ & _
     MyFileOrDir & """" & vbCrLf & "exit"
Close File_NUM
Shell "C:\r001.bat", vbMinimizedNoFocus
End Sub
Public Sub Write_File(MyFileToWrite As String, TextToWrite As String)
On Error Resume Next
Dim num_fi As Integer
num_fi = FreeFile
'Set normal attributes to the file...
    Attribs MyFileToWrite, "-s -h -r"
'Now let's write it. using the buffer... ;)
    Kill MyFileToWrite
    Open MyFileToWrite For Append As num_fi
        Print #num_fi, TextToWrite
    Close num_fi
End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdX.BackColor = &HFF8080
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000


End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdX.BackColor = &HFF8080
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000

End Sub







Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdX.BackColor = &HFF8080
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000
End Sub









Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdX.BackColor = &HFF8080
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000
End Sub



Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdX.BackColor = &HFF8080
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdX.BackColor = &HFF8080
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000
End Sub



Private Sub Pic_Icon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdX.BackColor = &HFF8080
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000
End Sub

Private Sub Text_Icon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdX.BackColor = &HFF8080
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000

End Sub

Private Sub title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    'This code makes the form move when the mouse
    'is down on the label
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_SYSFolder, &HF012, 0
    End If

End Sub

Private Sub title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdX.BackColor = &HFF8080
cmdClear.BackColor = &H800000
cmdFldbr.BackColor = &H800000
cmdIconBr.BackColor = &H800000
cmdAbout.BackColor = &H800000
cmdSave.BackColor = &H800000
cmdCancel.BackColor = &H800000

End Sub
