VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   3615
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   3495
   Icon            =   "frm_About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
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
      Height          =   1275
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frm_About.frx":08CA
      Top             =   960
      Width           =   3195
   End
   Begin FolderIcon.XLinkLabel XLinkLabel1 
      Height          =   255
      Left            =   840
      Top             =   2580
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NormTextColor   =   16761024
      HoverTextColor  =   16744576
      URL             =   "mailto:thor@eego.net?subject=FolderIcon"
      Caption         =   "thor@eego.net"
   End
   Begin VB.Label ICQ 
      BackColor       =   &H80000007&
      Caption         =   "15562140"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   840
      MouseIcon       =   "frm_About.frx":0AD2
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "E-mail :"
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
      Height          =   255
      Left            =   150
      TabIndex        =   7
      Top             =   2580
      Width           =   675
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Caption         =   "ICQ# :"
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
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   2880
      Width           =   555
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "For comments, suggestions or info :"
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
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "by Matthias ""Thor"" Stevens"
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
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   660
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "FolderIcon v1.0"
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
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   3480
      X2              =   3480
      Y1              =   5220
      Y2              =   -60
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   5280
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   -120
      X2              =   5820
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   7560
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About FolderIcon"
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
      TabIndex        =   1
      Top             =   0
      Width           =   3510
   End
   Begin VB.Label cmdClose 
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
      Height          =   270
      Left            =   1260
      TabIndex        =   0
      Top             =   3195
      Width           =   1005
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSFolder = &H112

Private Sub cmdClose_Click()
frm_Main.Enabled = True
Me.Hide
End Sub

Private Sub cmdClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &HFFC0C0
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &HFF8080
ICQ.ForeColor = &HFFC0C0
ICQ.FontUnderline = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H800000
ICQ.ForeColor = &HFFC0C0
ICQ.FontUnderline = False
End Sub

Private Sub ICQ_Click()
'create contact
    Dim fso, txtfile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtfile = fso.CreateTextFile("c:\contact.uin", True)
    txtfile.WriteLine ("")
    txtfile.WriteLine ("[ICQ Message User]")
    txtfile.WriteLine ("UIN=15562140")
    txtfile.WriteLine ("Email=thor@eego.net")
    txtfile.WriteLine ("NickName=Thor")
    txtfile.WriteLine ("FirstName=Matthias")
    txtfile.WriteLine ("LastName=Stevens")
    txtfile.Close

'launch contact
Call ShellExecute(0&, vbNullString, "C:\contact.uin", vbNullString, vbNullString, vbNormalFocus)

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H800000
ICQ.ForeColor = &HFFC0C0
ICQ.FontUnderline = False
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H800000
ICQ.ForeColor = &HFFC0C0
ICQ.FontUnderline = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H800000
ICQ.ForeColor = &HFFC0C0
ICQ.FontUnderline = False
End Sub

Private Sub ICQ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ICQ.ForeColor = &HFF8080
ICQ.FontUnderline = True
End Sub

Private Sub Text_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H800000
ICQ.ForeColor = &HFFC0C0
ICQ.FontUnderline = False
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
cmdClose.BackColor = &H800000
ICQ.ForeColor = &HFFC0C0
ICQ.FontUnderline = False
End Sub



