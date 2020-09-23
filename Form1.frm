VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Get Font Name Example"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Get Font Name"
      Height          =   435
      Left            =   4980
      TabIndex        =   3
      Top             =   1140
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   6660
      TabIndex        =   2
      Top             =   420
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   6555
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   1200
      Width           =   4755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Font Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Path To A Font File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'so we dont have to use commdlg ocx
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (opfn As OPENFILENAME) As Long

'used with getopenfilename api
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCusFilt(7) As Byte
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    lpFlags As Long
    nFiller(3) As Byte
    lpstrDefExt As String
    lCustData(11) As Byte
End Type

Private Sub Command1_Click()

    Dim udtOpenFile As OPENFILENAME
    Dim lngResult As Long

    With udtOpenFile
        .lStructSize = Len(udtOpenFile)
        .lpstrDefExt = "ttf"
        .lpstrFilter = "True Type Font" & vbNullChar & "*.ttf" & vbNullChar & vbNullChar
        .nMaxFile = 280
        .lpstrFile = String$(280, vbNullChar)
        lngResult = GetOpenFileName(udtOpenFile)
        If lngResult Then
            Text1 = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
        End If
    End With

End Sub

Private Sub Command2_Click()
    'retrieve the font name from the file
    Label2 = GetFontName(Text1)
End Sub

