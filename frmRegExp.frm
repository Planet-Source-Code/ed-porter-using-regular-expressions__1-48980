VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRegExp 
   Caption         =   "Regular Expression Test"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstPatterns 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      ItemData        =   "frmRegExp.frx":0000
      Left            =   2280
      List            =   "frmRegExp.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picShadow1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picmnu_Color 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   8280
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   1155
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   27
         Left            =   840
         Top             =   1560
         Width           =   225
         BackColor       =   8388608
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   26
         Left            =   600
         Top             =   1560
         Width           =   225
         BackColor       =   12582912
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   25
         Left            =   360
         Top             =   1560
         Width           =   225
         BackColor       =   16711680
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   24
         Left            =   120
         Top             =   1560
         Width           =   225
         BackColor       =   16744576
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   23
         Left            =   840
         Top             =   1320
         Width           =   225
         BackColor       =   8421376
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   22
         Left            =   600
         Top             =   1320
         Width           =   225
         BackColor       =   12632064
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   21
         Left            =   360
         Top             =   1320
         Width           =   225
         BackColor       =   16776960
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   20
         Left            =   120
         Top             =   1320
         Width           =   225
         BackColor       =   16777152
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   19
         Left            =   840
         Top             =   1080
         Width           =   225
         BackColor       =   32768
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   18
         Left            =   600
         Top             =   1080
         Width           =   225
         BackColor       =   49152
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   17
         Left            =   360
         Top             =   1080
         Width           =   225
         BackColor       =   65280
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   16
         Left            =   120
         Top             =   1080
         Width           =   225
         BackColor       =   8454016
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   15
         Left            =   840
         Top             =   840
         Width           =   225
         BackColor       =   32896
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   14
         Left            =   600
         Top             =   840
         Width           =   225
         BackColor       =   49344
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   13
         Left            =   360
         Top             =   840
         Width           =   225
         BackColor       =   65535
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   12
         Left            =   120
         Top             =   840
         Width           =   225
         BackColor       =   12648447
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   11
         Left            =   840
         Top             =   600
         Width           =   225
         BackColor       =   16512
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   10
         Left            =   600
         Top             =   600
         Width           =   225
         BackColor       =   16576
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   9
         Left            =   360
         Top             =   600
         Width           =   225
         BackColor       =   33023
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   8
         Left            =   120
         Top             =   600
         Width           =   225
         BackColor       =   8438015
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   7
         Left            =   840
         Top             =   360
         Width           =   225
         BackColor       =   128
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   6
         Left            =   600
         Top             =   360
         Width           =   225
         BackColor       =   4210816
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   5
         Left            =   360
         Top             =   360
         Width           =   225
         BackColor       =   255
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   4
         Left            =   120
         Top             =   360
         Width           =   225
         BackColor       =   8421631
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   3
         Left            =   840
         Top             =   120
         Width           =   225
         BackColor       =   0
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   2
         Left            =   600
         Top             =   120
         Width           =   225
         BackColor       =   8421504
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   1
         Left            =   360
         Top             =   120
         Width           =   225
         BackColor       =   12583104
         BorderStyle     =   0
         Size            =   "388;388"
      End
      Begin MSForms.Image imgColorMenu 
         Height          =   225
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   225
         BackColor       =   16777215
         BorderStyle     =   0
         Size            =   "388;388"
      End
   End
   Begin VB.ListBox lstFontSize 
      Appearance      =   0  'Flat
      Height          =   615
      ItemData        =   "frmRegExp.frx":0018
      Left            =   7320
      List            =   "frmRegExp.frx":001F
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstFontName 
      Appearance      =   0  'Flat
      Height          =   420
      ItemData        =   "frmRegExp.frx":0030
      Left            =   5400
      List            =   "frmRegExp.frx":0037
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picmnu_Properties 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   4
      Top             =   570
      Visible         =   0   'False
      Width           =   4215
      Begin MSForms.ToggleButton tbFontSize 
         Height          =   225
         Left            =   3060
         TabIndex        =   10
         ToolTipText     =   "Select Font Size"
         Top             =   30
         Width           =   225
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "388;388"
         Value           =   "0"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblFontSize 
         Height          =   195
         Left            =   2730
         TabIndex        =   9
         Top             =   30
         Width           =   300
         BackColor       =   -2147483639
         VariousPropertyBits=   8388627
         Caption         =   "10"
         Size            =   "529;344"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton tbFontName 
         Height          =   225
         Left            =   2490
         TabIndex        =   8
         ToolTipText     =   "Select Font Name"
         Top             =   30
         Width           =   225
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "388;388"
         Value           =   "0"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lblFontName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Times New Roman"
         Height          =   195
         Left            =   1050
         TabIndex        =   7
         Top             =   30
         Width           =   1350
      End
      Begin MSForms.ToggleButton tbFontColor 
         Height          =   225
         Left            =   3795
         TabIndex        =   6
         ToolTipText     =   "Select Font Color"
         Top             =   30
         Width           =   225
         BackColor       =   16711680
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "397;397"
         Value           =   "0"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lblFontColor 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Left            =   3360
         TabIndex        =   5
         Top             =   30
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Clear"
      Height          =   360
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtPattern 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin RichTextLib.RichTextBox rtbTest 
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1720
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmRegExp.frx":0048
   End
   Begin VB.Label lblMeasure 
      AutoSize        =   -1  'True
      Caption         =   "lblMeasure"
      Height          =   195
      Left            =   3840
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   765
   End
   Begin MSForms.ToggleButton tbPatterns 
      Height          =   360
      Left            =   2280
      TabIndex        =   18
      ToolTipText     =   "Select Font Name"
      Top             =   120
      Width           =   975
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1720;635"
      Value           =   "0"
      Caption         =   "Patterns"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "lblTime"
      Height          =   195
      Left            =   1680
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblMatchesFound 
      AutoSize        =   -1  'True
      Caption         =   "lblMatchesFound"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ClientToScreen Lib "user32" _
(ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long

Private Declare Function LBItemFromPt Lib "comctl32.dll" _
(ByVal hLB As Long, ByVal ptX As Long, ByVal ptY As Long, _
ByVal bAutoScroll As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Any) As Long   ' <---

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Type POINTAPI   ' pt
  X As Long
  Y As Long
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Const WM_SETREDRAW = &HB

Dim flgTest As Boolean

Private Sub LoadPatterns()
'Load lstPatterns with RegExp patterns
Dim strItem As String
Dim I As Long, J As Long

With lstPatterns
    ' ("[^"]*")
    .Clear
    strItem = "(" + Chr(34) + Chr(91) + "^" + Chr(34) + Chr(93) + "*" + _
                    Chr(34) + ")" + "    Find all characters enclosed by quotes"
    .AddItem strItem
    strItem = "('[^\r]*)(\r\n)?    Highlight all comment lines"
    .AddItem strItem
    strItem = "Dim[^\w]|Function[^\w]" + "    Group espressions using " + _
                                                    Chr(34) + Chr(124) + Chr(34)
    .AddItem strItem
    strItem = "Public[^\w]|Private[^\w]"
    .AddItem strItem
    strItem = "On[^\w]|Exit|End|Sub[^\w]|True[^\w]|False[^\w]|For[^\w]| Next"
    .AddItem strItem
    
    .Top = tbPatterns.Top + tbPatterns.Height + 2
    .Left = tbPatterns.Left
    .Height = .ListCount * 20
    'Find listbox width
    lblMeasure.FontSize = .FontSize
    lblMeasure.FontName = .FontName
    For I = 0 To .ListCount - 1
        lblMeasure.Caption = .List(I)
        If J < lblMeasure.Width + 12 Then J = lblMeasure.Width + 25
    Next I
    .Width = J
End With

End Sub

Private Sub cmdTest_Click(Index As Integer)
Dim sysTime As SYSTEMTIME
Dim tA As Long, tB As Long, ElapsedTime As Single

Me.MousePointer = vbHourglass
With rtbTest
    SendMessage .hWnd, WM_SETREDRAW, False, 0
    flgTest = True
    
    Select Case Index
    Case 0 'Test script
        '--------------------------- Time function --------------------------
        GetSystemTime sysTime
        With sysTime
            tA = Val(.wMinute) * 60000 + Val(.wSecond) * 1000 + Val(.wMilliseconds)
        End With
        '---------------------------------------------------------------------

        ColorSyntax rtbTest, txtPattern.Text
        
        '--------------------------- Time function --------------------------
        GetSystemTime sysTime
        With sysTime
            tB = Val(.wMinute) * 60000 + Val(.wSecond) * 1000 + Val(.wMilliseconds)
        End With
        ElapsedTime = (tB - tA) / 1000 'In seconds
        With lblTime
            .Caption = "Time to Format Code: " + Str(ElapsedTime) + " secs."
            .Left = lblMatchesFound.Left + lblMatchesFound.Width + 12
            .Visible = True
            .ZOrder
        End With
        '---------------------------------------------------------------------
        .SelStart = 0

    Case 1 'Reset
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = vbBlack
        .SelFontName = "Ms Sans Serif"
        .SelFontSize = 10
        .SelStart = 0
        .Refresh
        lblMatchesFound.Visible = False
        lblTime.Visible = False
    End Select
    
    flgTest = False
    SendMessage .hWnd, WM_SETREDRAW, True, 0
    .Refresh
End With

Me.MousePointer = vbArrow

End Sub

Public Sub ColorSyntax(ByRef objRTB As RichTextBox, Pattern As String)
' To use add 'Reference' to 'Microsoft VBScript Regular Expressions vx.x
Dim SearchString As String
Dim nRegExp As New RegExp
Dim Matches As MatchCollection
Dim CharMatch As Match
Dim I As Long

If Pattern = "" Then Exit Sub 'Prevents latchup

'Tested Patterns
' ("[^"]*")   Find all characters enclosed by "..."
' ('[^\r]*)(\r\n)?    Highlight all lines starting with "'" (ie. comments)

'Instructions - vbBlue
' Dim[^\w]
' Function[^\w]
' Public[^\w]
' Private[^\w]
' On[^\w]
' Exit
' End
' Sub[^\w]
' True[^\w]
' False[^\w]
' For[^\w]
' Next
' If
' ElseIf
' Then


'nRegExp.Test
SearchString = Pattern
    
    'Parse objRTB using RegExp pattern
    nRegExp.Pattern = SearchString
    nRegExp.Global = True
    'Load the match collection with pointers to the instances within the
    'richtextbox control
    Set Matches = nRegExp.Execute(objRTB.Text)
    'Loop on the collection changing the color in the richtextbox control
    For Each CharMatch In Matches
        'Use the select method and the selection property of the
        'richtextbox control
        With objRTB
            .SelStart = CharMatch.FirstIndex
            .SelLength = CharMatch.Length
            .SelColor = tbFontColor.BackColor
            'Don't use Black as a test color
            If .SelColor = vbBlack Then .SelColor = vbRed
            .SelFontName = lblFontName.Caption
            .SelFontSize = Val(lblFontSize.Caption)
        End With
        '**************************************
        'Count the Number found
        I = I + 1
        '**************************************
    Next
    With lblMatchesFound
        .Caption = Str(I) + " - Matches Found"
        .Visible = True
        .ZOrder
    End With
    
    objRTB.SelLength = 0
    'Get rid of the objects
    Set nRegExp = Nothing
    Set Matches = Nothing
    Set CharMatch = Nothing

End Sub


Private Sub Form_Load()
Dim I As Long
Dim strFile As String

With Me
    .Top = 0
    .Left = 0
    .Width = Screen.Width
    .Height = Screen.Height - 400
    .ScaleMode = vbPixels
End With

With cmdTest(0)
    .Top = 8
    .Left = 8
    .Width = 65
    .Height = 24
    .Visible = True
    .ZOrder
End With
With cmdTest(1)
    .Top = cmdTest(0).Top
    .Left = 80
    .Width = cmdTest(0).Width
    .Height = cmdTest(0).Height
    .Visible = True
    .ZOrder
End With
With tbPatterns
    .Top = cmdTest(0).Top
    .Left = 152
    .Width = cmdTest(0).Width
    .Height = cmdTest(0).Height
    .Visible = True
    .ZOrder
End With

With txtPattern
    .Left = tbPatterns.Left + tbPatterns.Width + 4
    .Top = cmdTest(0).Top
    .Height = cmdTest(0).Height
    .Width = Me.ScaleWidth - .Left - cmdTest(0).Left
End With

With rtbTest
    .Top = 64
    .Left = cmdTest(0).Left
    .Width = Me.ScaleWidth - .Left - cmdTest(0).Left
    .Height = Me.ScaleHeight - .Top - cmdTest(0).Top
End With

'Load Font Name(s) info into listbox
With lstFontName
    .Clear
    For I = 0 To Screen.FontCount - 1
      .AddItem Screen.Fonts(I)
    Next I
    .Height = .ListCount * 14
    .ListIndex = 0
    .ZOrder
End With

'Load Font Size(s) into listbox
With lstFontSize
    .Clear
    For I = 8 To 12
        .AddItem I
    Next I
    For I = 14 To 48 Step 2
        If I = 26 Then I = I + 8
        If I = 36 Then I = I + 12
        .AddItem I
    Next I
    .Height = .ListCount * 14
    .ListIndex = 0
    .ZOrder
End With

LoadPatterns 'Load Reg Exp test patterns
FormatMnuProperties

'Load test script
strFile = App.Path + "\Test.rtf"
With rtbTest
    .LoadFile strFile
    .SelStart = 0
    .SelLength = Len(.TextRTF)
    .SelIndent = 10
    .SelRightIndent = 10
    .SelLength = 0
End With
    
lblMatchesFound.Caption = ""

End Sub
Private Sub FormatMnuProperties()

With picmnu_Properties
    .Width = lblFontName.Width + tbFontName.Width _
                + lblFontSize.Width + tbFontSize.Width + lblFontColor.Width _
                + tbFontColor.Width + 28
    .Left = Me.ScaleWidth - .Width - cmdTest(0).Left + 4
    .Visible = True
    .ZOrder
End With

With tbFontColor
    .Left = picmnu_Properties.Width - .Width - 4
End With

With lblFontColor
    .Left = tbFontColor.Left - .Width - 4
End With
    
With tbFontSize
    .Left = lblFontColor.Left - .Width - 4
End With

With lblFontSize
    .Left = tbFontSize.Left - .Width - 4
End With

With tbFontName
    .Left = lblFontSize.Left - .Width - 4
End With

With lblFontName
    .Left = tbFontName.Left - .Width - 4
End With

End Sub


Private Sub imgColorMenu_Click(Index As Integer)

With tbFontColor
    .BackColor = imgColorMenu(Index).BackColor
    .BackStyle = fmBackStyleOpaque
End With

tbFontColor.Value = False

End Sub

Private Sub lstFontName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ItemIndex As Long
Dim lbPoint As POINTAPI

lbPoint.X = X \ Screen.TwipsPerPixelX
lbPoint.Y = Y \ Screen.TwipsPerPixelY
With lstFontName
    Call ClientToScreen(.hWnd, lbPoint)
    ItemIndex = LBItemFromPt(.hWnd, lbPoint.X, lbPoint.Y, False)
    If .ListCount > 0 Then .ListIndex = ItemIndex
End With

End Sub


Private Sub lstFontName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With lstFontName
    lblFontName.Caption = .List(.ListIndex)
End With
tbFontName.Value = False
FormatMnuProperties

End Sub


Private Sub lstFontSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ItemIndex As Long
Dim lbPoint As POINTAPI

lbPoint.X = X \ Screen.TwipsPerPixelX
lbPoint.Y = Y \ Screen.TwipsPerPixelY
With lstFontSize
    Call ClientToScreen(.hWnd, lbPoint)
    ItemIndex = LBItemFromPt(.hWnd, lbPoint.X, lbPoint.Y, False)
    If .ListCount > 0 Then .ListIndex = ItemIndex
End With

End Sub


Private Sub lstFontSize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With lstFontSize
        lblFontSize.Caption = .List(.ListIndex)
        tbFontSize.Value = False
End With

End Sub


Private Sub lstPatterns_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ItemIndex As Long
Dim lbPoint As POINTAPI

lbPoint.X = X \ Screen.TwipsPerPixelX
lbPoint.Y = Y \ Screen.TwipsPerPixelY
With lstPatterns
    Call ClientToScreen(.hWnd, lbPoint)
    ItemIndex = LBItemFromPt(.hWnd, lbPoint.X, lbPoint.Y, False)
    If .ListCount > 0 Then .ListIndex = ItemIndex
End With

End Sub


Private Sub lstPatterns_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With lstPatterns
    
    If .ListIndex = 0 Then
        txtPattern.Text = "(" + Chr(34) + Chr(91) + "^" + Chr(34) + Chr(93) + _
                                        "*" + Chr(34) + ")"
    ElseIf .ListIndex = 1 Then
        txtPattern.Text = "('[^\r]*)(\r\n)? "
    ElseIf .ListIndex = 2 Then
        txtPattern.Text = "Dim[^\w]|Function[^\w]"
    Else
        txtPattern.Text = .List(.ListIndex)
    End If
End With
tbPatterns.Value = False

End Sub


Private Sub rtbTest_SelChange()

Exit Sub
If flgTest = True Then Exit Sub

With rtbTest
    If Not IsNull(.SelFontName) Then lblFontName.Caption = .SelFontName
    If Not IsNull(.SelFontSize) Then lblFontSize.Caption = .SelFontSize
    If Not IsNull(.SelColor) Then tbFontColor.BackColor = .SelColor
End With
    
End Sub



Private Sub tbFontColor_Click()

If tbFontColor.Value = True Then
        With picmnu_Color
            .Left = picmnu_Properties.Left + picmnu_Properties.Width - .Width + 4
            .Top = picmnu_Properties.Top + picmnu_Properties.Height + 2
            .Visible = True
            .ZOrder
        End With
    Else
        picmnu_Color.Visible = False
        picShadow1.Visible = False
End If
tbFontColor.BackStyle = fmBackStyleOpaque

End Sub


Private Sub tbFontName_Click()
Dim I As Long

If tbFontName.Value = True Then

    With lstFontName
        .Left = picmnu_Properties.Left + lblFontName.Left
        .Top = picmnu_Properties.Top + picmnu_Properties.Height
        'Cycle through list to find index of font name
        For I = 0 To .ListCount - 1
            If .List(I) = lblFontName.Caption Then
                .ListIndex = I
                Exit For
            Else
                .ListIndex = 0
            End If
        Next I
        .Visible = True
        .ZOrder
    End With
Else
    lstFontName.Visible = False
End If

End Sub

Private Sub tbFontSize_Click()
Dim I As Integer

If tbFontSize.Value = True Then

    With lstFontSize
        .Left = picmnu_Properties.Left + lblFontSize.Left
        .Top = picmnu_Properties.Top + picmnu_Properties.Height
        'Cycle through list to find correct index
        For I = 0 To .ListCount - 1
            If .List(I) = lblFontSize.Caption Then
                .ListIndex = I
                Exit For
            Else
                .ListIndex = 0
            End If
        Next I
        .Visible = True
        .ZOrder
    End With

ElseIf tbFontSize.Value = False Then
    lstFontSize.Visible = False
End If

End Sub


Private Sub tbPatterns_Click()

With tbPatterns
    If .Value = True Then
        lstPatterns.Visible = True
        lstPatterns.ZOrder
    Else
        lstPatterns.Visible = False
    End If
End With

End Sub


