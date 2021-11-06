VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Extract"
   ClientHeight    =   2355
   ClientLeft      =   2580
   ClientTop       =   2295
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "ICONEXTR.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   4170
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   1800
      Width           =   492
   End
   Begin VB.CommandButton btnCopy 
      Appearance      =   0  'Flat
      Caption         =   "Copy to picture box ->"
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2052
   End
   Begin VB.HScrollBar hs 
      Enabled         =   0   'False
      Height          =   252
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      Height          =   372
      Left            =   3000
      TabIndex        =   7
      Top             =   1080
      Width           =   1092
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   840
      Width           =   492
   End
   Begin VB.CommandButton btnAbout 
      Appearance      =   0  'Flat
      Caption         =   "About..."
      Height          =   372
      Left            =   3000
      TabIndex        =   8
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton btnOpen 
      Appearance      =   0  'Flat
      Caption         =   "Open..."
      Height          =   372
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label lblNumIcons 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Width           =   612
   End
   Begin VB.Label lblDumb 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Number of Icons in file:"
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iconn% 'The current icon number in a file
Dim iconfilename$ 'The filename of the icon file(.EXE, .DLL, .ICO)
Dim numicons% 'The number of icons in a file

Private Declare Function DrawIcon Lib "user" (ByVal hDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal hIcon As Integer) As Integer
'hDC- Device context of the control to be drawn to
'x, y- coordinates of where to draw the icon in the control
'hIcon-Handle of an icon

Private Declare Function ExtractIcon Lib "shell.dll" (ByVal hinst%, ByVal lpszExeName$, ByVal iIcon%) As Integer
'hinst- The instance handle of the application calling ExtractIcon.  Should be the name of your EXE file, or VB.EXE at runtime
'lpszExeName- Module containing icons
'iIcon%- number of the icon in the file.  If you put -1 for this, it returns the amount of icons in a file
'The return value should be: 1)An icon handle 2)1 if it's not a EXE, DLL, or ICO file 3)NULL if no icons are in a file

Private Declare Function GetModuleHandle Lib "Kernel" (ByVal lpModuleName As String) As Integer
'lpModuleName- The filename of a module, to get the handle of it.

Private Sub AboutProg()

End Sub

Private Sub btnAbout_Click()
nl$ = Chr$(13) + Chr$(10)
Msg$ = "Icon Extract, Written By Tim Hill (AOL: Tim112)" + nl$ + "Feel free to use this code in your programs." + nl$ + "Tim Hill shall not be responsible in any way for any type of damages incurred to you or your system during the use of this software."
MsgBox Msg$, 64, "Icon Extract"
End Sub

Private Sub btnCopy_Click()
pic2.Picture = pic.Image 'Must be pic2.Picture = pic.IMAGE, not pic.Picture, because it is not
'actually part of the picture yet when you use the API call
End Sub

Private Sub btnOpen_Click()
iconfilename$ = InputBox$("Icon File(.ICO,.EXE,.DLL):", "Icon Extract", "PROGMAN.EXE") 'Asks for filename
If iconfilename$ = "" Then Exit Sub
pic.Cls 'clears the picture box
hModule = GetModuleHandle("C:\VB.EXE") 'gets handle
iconmod$ = iconfilename$ + Chr$(0) 'prepares filename
Iconh = ExtractIcon(hModule, iconmod$, -1) 'gets number of icons
numicons% = Iconh 'puts it into a variable
lblNumIcons.Caption = Str$(numicons%) 'shows number of icons on label
numicons% = numicons% - 1 'Accounts for the first icon, at number 0
If numicons% > 1 Then 'disables scroll bar if only one or less
hs.Enabled = -1
Else
hs.Enabled = 0
End If
Iconh = ExtractIcon(hModule, iconmod$, 0) 'Extracts the first icon
x% = DrawIcon(pic.hDC, 0, 0, Iconh) 'Draws the first icon
hs.Max = numicons% 'sets maximum scroll bar value to the number of icons
hs.Value = 0
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
AboutProg
End Sub

Private Sub hs_Change()
pic.Cls 'Clears the picture box
iconn% = hs.Value 'sets the value of the icon number to the scroll bar position
hModule = GetModuleHandle("C:\VB.EXE") 'Gets the module handle
iconmod$ = iconfilename$ + Chr$(0) 'prepares filename for ExtractIcon
Iconh = ExtractIcon(hModule, iconmod$, iconn%) 'Extracts the specified icon
x% = DrawIcon(pic.hDC, 0, 0, Iconh) 'Draws icon
End Sub

Private Sub up_Click()
pic.Cls
iconn% = iconn% + 1
If iconn% > dori% Then iconn% = dori%
hModule = GetModuleHandle("C:\VB.EXE")
iconmod$ = "C:\WINDOWS\PROGMAN.EXE" + Chr$(0)
Iconh = ExtractIcon(hModule, iconmod$, iconn%)
x% = DrawIcon(pic.hDC, 0, 0, Iconh)
End Sub

