VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Images-in-a-Cube"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   428
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Moldura 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   6555
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   17
      Top             =   4755
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox Picture2 
      Height          =   285
      Left            =   5370
      MouseIcon       =   "Main.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   375
      TabIndex        =   15
      Top             =   4365
      Width           =   435
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Use Frame"
      Height          =   300
      Left            =   3855
      TabIndex        =   13
      Top             =   4020
      Width           =   2130
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Cube at Start up"
      Height          =   300
      Left            =   3855
      TabIndex        =   12
      Top             =   3705
      Width           =   2130
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   6870
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   11
      Top             =   3180
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Cube"
      Height          =   885
      Left            =   4230
      TabIndex        =   10
      Top             =   6975
      Width           =   1275
   End
   Begin VB.PictureBox original 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1575
      Left            =   7065
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   30
      Pattern         =   "*.gif;*.bmp;*.jpg"
      TabIndex        =   2
      Top             =   6600
      Width           =   3210
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   30
      TabIndex        =   1
      Top             =   5580
      Width           =   3225
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   5205
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Transparent color"
      Height          =   195
      Left            =   3855
      TabIndex        =   16
      Top             =   4365
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Frame"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   6
      Left            =   4095
      TabIndex        =   14
      Top             =   4710
      Width           =   1500
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1680
      Index           =   6
      Left            =   4080
      MouseIcon       =   "Main.frx":030A
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Index           =   5
      Left            =   2280
      TabIndex        =   9
      Top             =   4035
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Index           =   4
      Left            =   5325
      TabIndex        =   8
      Top             =   2250
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Index           =   3
      Left            =   3795
      TabIndex        =   7
      Top             =   2250
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Index           =   2
      Left            =   2280
      TabIndex        =   6
      Top             =   2250
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Index           =   1
      Left            =   675
      TabIndex        =   5
      Top             =   2250
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   645
      Width           =   330
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1680
      Index           =   5
      Left            =   1665
      Stretch         =   -1  'True
      Top             =   3435
      Width           =   1575
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1680
      Index           =   0
      Left            =   1665
      Stretch         =   -1  'True
      Top             =   105
      Width           =   1575
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1680
      Index           =   4
      Left            =   4785
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   1575
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1680
      Index           =   3
      Left            =   3225
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   1575
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1680
      Index           =   2
      Left            =   1665
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   1575
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1680
      Index           =   1
      Left            =   105
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Private Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "GDI32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long

Private Const STRETCHMODE As Long = vbPaletteModeNone

Private Escolha As Integer
Private Image_files(0 To 6) As String
Private Cor As Long


Private Sub Command1_Click()

  Dim i As Integer
    
    Me.Hide

    Call SetStretchBltMode(Picture1.hDC, STRETCHMODE)
    For i = 0 To 5
        original.Picture = Image1(i).Picture
        StretchBlt Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, original.hDC, 0, 0, original.ScaleWidth, original.ScaleHeight, vbSrcCopy
        Picture1.Refresh
        If Check2 Then
            original.Picture = Image1(6).Picture
            GdiTransparentBlt Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, original.hDC, 0, 0, original.ScaleWidth, original.ScaleHeight, Picture2.BackColor
        End If
        
         Picture1.Refresh
        'Remove_transColor
        SavePicture Picture1.Image, App.Path & "\Side" & Str(i + 1) & ".bmp"

        SaveSetting "Images-in-a-Cube", "Image", Str(i), Image_files(i)

    Next i
    
    SaveSetting "Images-in-a-Cube", "Image", Str(6), Image_files(i)
    SaveSetting "Images-in-a-Cube", "Startup", "Value", Check1.Value
    SaveSetting "Images-in-a-Cube", "Tranarent color", "Value", Picture2.BackColor
    SaveSetting "Images-in-a-Cube", "Use Frame", "Value", Check2.Value
    
    DoEvents
    frmMain.Show

    If Terminou Then
        Unload frmMain
        Me.Show
        Terminou = False
    End If

End Sub

Private Sub Dir1_Change()

    File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()

    original.Picture = LoadPicture(Dir1.Path & "\" & File1.List(File1.ListIndex))

    Image1(Escolha).Picture = LoadPicture(Dir1.Path & "\" & File1.List(File1.ListIndex))
    
    Image_files(Escolha) = Dir1.Path & "\" & File1.List(File1.ListIndex)
    If Escolha = 6 Then
    Moldura.Picture = LoadPicture(Dir1.Path & "\" & File1.List(File1.ListIndex))
    End If
    
End Sub

Private Sub Form_Load()
    'On Error GoTo erro
    
    Dim i As Integer
    
    Drive1.Drive = Left$(App.Path, 3)
    Dir1.Path = App.Path
    
    For i = 0 To 6
        Image_files(i) = GetSetting("Images-in-a-Cube", "Image", Str(i), "")
        Image1(i).Picture = LoadPicture(Image_files(i))
    Next
        Moldura.Picture = LoadPicture(Image_files(6))
        
    Check1.Value = GetSetting("Images-in-a-Cube", "Startup", "Value", 0)
    
    Check2.Value = GetSetting("Images-in-a-Cube", "Use Frame", "Value", 0)
   
    Picture2.BackColor = GetSetting("Images-in-a-Cube", "Tranarent color", "Value", &HFF)
    
    If Check1.Value Then Command1_Click
    
sair:
    Exit Sub
    
erro:
    Image_files(i) = ""
    Resume Next
    
    
End Sub

Private Sub Image1_Click(Index As Integer)

  Static Anterior As Integer

    If Index <> Escolha Then
        Label1(Anterior).Enabled = False
        Escolha = Index
        Label1(Index).Enabled = True
        Anterior = Index
    End If

If Index = 6 Then
    If Image1(6).MousePointer = 99 Then
        Image1(6).MousePointer = 0
        Picture2.BackColor = Cor
        Picture2.MousePointer = 0
    End If
End If

End Sub

Private Sub Remove_transColor()
    Dim i As Integer
    Dim k As Integer
    
    For i = 0 To Picture1.ScaleWidth
        For k = 0 To Picture1.ScaleHeight
            If Picture1.Point(i, k) = 0 Then
                Picture1.PSet (i, k), 1
            End If
        Next k
    Next i
    Picture1.Refresh

End Sub


Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 6 And Image1(6).MousePointer = 99 Then
    X = X / Screen.TwipsPerPixelX
    Y = Y / Screen.TwipsPerPixelY
    X = Moldura.ScaleWidth * X / Image1(6).Width
    Y = Moldura.ScaleHeight * Y / Image1(6).Width
    Cor = Abs(Moldura.Point(X, Y))
    Picture2.BackColor = Cor
End If

End Sub

Private Sub Picture2_Click()
Picture2.MousePointer = 99
Image1(6).MousePointer = 99
End Sub
