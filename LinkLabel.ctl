VERSION 5.00
Begin VB.UserControl XLinkLabel 
   CanGetFocus     =   0   'False
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   MouseIcon       =   "LinkLabel.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   390
   ScaleWidth      =   420
End
Attribute VB_Name = "XLinkLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' XLinkLabel
' Author: David Crowell (davidc@qtm.net)
' See http://www.qtm.net/~davidc for updates
' Released to the public domain
'
' Last update: May 22, 1999
'
' Use this code at your own risk.  I assume no liability for
' the use of this code.
'
' The purpose of XLinkLabel is to have a label control
' that works as a hyperlink.
'

' API Declares
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

' Module level variables
Private nControlHeight As Long
Private nControlWidth As Long
Private bHovering As Boolean

' Property member variables
Private mBackColor As OLE_COLOR
Private mNormTextColor As OLE_COLOR
Private mHoverTextColor As OLE_COLOR
Private mNormUnderline As Boolean
Private mHoverUnderline As Boolean
Private mFont As StdFont
Private mCaption As String
Private mURL As String
Private mEnabled As Boolean

' The property Get/Let/Set stuff if pretty
' self-explanatory, so figure it out :)

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = mBackColor
End Property
Public Property Let BackColor(NewColor As OLE_COLOR)
    mBackColor = NewColor
    UserControl.BackColor = mBackColor
    UserControl_Paint
    PropertyChanged "BackColor"
End Property

Public Property Get NormTextColor() As OLE_COLOR
Attribute NormTextColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute NormTextColor.VB_UserMemId = -513
    NormTextColor = mNormTextColor
End Property
Public Property Let NormTextColor(NewColor As OLE_COLOR)
    mNormTextColor = NewColor
    UserControl.ForeColor = NewColor
    UserControl_Paint
    PropertyChanged "NormTextColor"
End Property

Public Property Get HoverTextColor() As OLE_COLOR
Attribute HoverTextColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HoverTextColor = mHoverTextColor
End Property
Public Property Let HoverTextColor(NewColor As OLE_COLOR)
    mHoverTextColor = NewColor
    UserControl_Paint
    PropertyChanged "HoverTextColor"
End Property

Public Property Get NormUnderline() As Boolean
    NormUnderline = mNormUnderline
End Property
Public Property Let NormUnderline(val As Boolean)
    mNormUnderline = val
    UserControl.FontUnderline = val
    UserControl_Paint
    PropertyChanged "NormUnderline"
End Property

Public Property Get HoverUnderline() As Boolean
Attribute HoverUnderline.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HoverUnderline = mHoverUnderline
End Property
Public Property Let HoverUnderline(val As Boolean)
    mHoverUnderline = val
    UserControl_Paint
    PropertyChanged "HoverUnderline"
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
    Set Font = mFont
End Property
Public Property Set Font(NewFont As StdFont)
    Set mFont = NewFont
    Set UserControl.Font = mFont
    UserControl_Paint
    PropertyChanged "Font"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    Caption = mCaption
End Property
Public Property Let Caption(val As String)
    mCaption = val
    UserControl_Paint
    PropertyChanged "Caption"
End Property

Public Property Get URL() As String
Attribute URL.VB_ProcData.VB_Invoke_Property = ";Behavior"
    URL = mURL
End Property
Public Property Let URL(val As String)
    mURL = val
    PropertyChanged "URL"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = mEnabled
End Property
Public Property Let Enabled(val As Boolean)
    mEnabled = val
    UserControl.Enabled = val
    PropertyChanged "Enabled"
End Property

' set up the default values
Private Sub UserControl_InitProperties()
    Set Font = Ambient.Font
    BackColor = Ambient.BackColor
    NormTextColor = Ambient.ForeColor
    HoverTextColor = Ambient.ForeColor
    URL = "http://www.qtm.net/~davidc" ' got to put my plug in here :)
    Enabled = True
    Caption = UserControl.Extender.Name
    NormUnderline = False
    HoverUnderline = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
' Read saved properties
    On Error Resume Next
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    NormTextColor = PropBag.ReadProperty("NormTextColor", Ambient.ForeColor)
    HoverTextColor = PropBag.ReadProperty("HoverTextColor", Ambient.ForeColor)
    URL = PropBag.ReadProperty("URL", "http://www.qtm.net/~davidc")
    Enabled = PropBag.ReadProperty("Enabled", True)
    Caption = PropBag.ReadProperty("Caption", "URL Label")
    NormUnderline = PropBag.ReadProperty("NormUnderline", False)
    HoverUnderline = PropBag.ReadProperty("HoverUnderline", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
' Write properties (used only during design time)
    PropBag.WriteProperty "Font", Font
    PropBag.WriteProperty "BackColor", BackColor, Ambient.BackColor
    PropBag.WriteProperty "NormTextColor", NormTextColor, Ambient.ForeColor
    PropBag.WriteProperty "HoverTextColor", HoverTextColor, Ambient.ForeColor
    PropBag.WriteProperty "URL", URL, "http://www.qtm.net/~davidc"
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "Caption", Caption, "URL Label"
    PropBag.WriteProperty "NormUnderline", NormUnderline, False
    PropBag.WriteProperty "HoverUnderline", HoverUnderline, True
End Sub

Private Sub UserControl_Resize()

    ' Store the size of the usercontrol for faster use later
    nControlHeight = UserControl.Height
    nControlWidth = UserControl.Width
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim bHover As Boolean
    
    ' Always call release capture
    Call ReleaseCapture
    
    ' Is the mouse over the control?
    If (x < 0) Or (y < 0) Or (x > nControlWidth) Or (y > nControlHeight) Then
        bHover = False
    Else
        bHover = True
        ' if so be sure to call SetCapture, so we'll catch it when it moves off
        Call SetCapture(UserControl.hWnd)
    End If
    
    If bHovering <> bHover Then
    ' only change appearance if necessary
        bHovering = bHover
        
        ' we change the font setting of the usercontrol
        ' which will take effect next time we repaint
        If bHovering Then
            UserControl.FontUnderline = mHoverUnderline
            UserControl.ForeColor = mHoverTextColor
        Else
            UserControl.FontUnderline = mNormUnderline
            UserControl.ForeColor = mNormTextColor
        End If
        
        ' repaint
        UserControl_Paint
        
    End If
    
End Sub

Private Sub UserControl_Paint()

    ' erase current contents
    UserControl.Cls
    
    ' and repaint the text
    Call TextOut(UserControl.hDC, 0, 0, mCaption, Len(mCaption))
    
End Sub

Private Sub UserControl_Click()
' Open the default browser with the URL

    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
    
End Sub

