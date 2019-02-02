VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMinesweeper 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mineweeper"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frtiv 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Image img1 
         Height          =   375
         Left            =   240
         Picture         =   "frmgame.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lbltiv 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   720
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.CommandButton btnexit 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Picture         =   "frmgame.frx":064A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   495
   End
   Begin VB.Timer tmrtime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   240
   End
   Begin MSComctlLib.ImageList imglst1 
      Left            =   360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   38
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgame.frx":0CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgame.frx":13FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnstart 
      Caption         =   "Start"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   6960
      Width           =   1215
   End
   Begin VB.PictureBox piccomp 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   360
      Picture         =   "frmgame.frx":1A54
      ScaleHeight     =   6060.606
      ScaleMode       =   0  'User
      ScaleWidth      =   6000
      TabIndex        =   1
      Top             =   720
      Width           =   6000
      Begin VB.Frame frmin 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   1680
         TabIndex        =   5
         Top             =   2880
         Width           =   2415
         Begin VB.ComboBox cmbmin 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            ItemData        =   "frmgame.frx":13CE4
            Left            =   360
            List            =   "frmgame.frx":13CE6
            TabIndex        =   7
            Text            =   "3"
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblmin 
            AutoSize        =   -1  'True
            Caption         =   "How long?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   600
            TabIndex        =   6
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.CommandButton btncomp 
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   100
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.Label lblardyunq 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2640
      TabIndex        =   11
      Top             =   120
      Width           =   90
   End
   Begin VB.Label lblcoopyright 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   7440
      Width           =   45
   End
   Begin VB.Label lbltime 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frmMinesweeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X(0 To 11, 0 To 11) As Integer, v As Integer, r As Integer, icomp As Integer, o As Integer, zs As Integer, ii As Integer, k As Boolean, cm(0 To 11, 0 To 11), i As Integer, j As Integer, iindex As Integer

Private Sub comp()
    icomp = 22
    Do
        comptest
    Loop While icomp > 0
    icomp = 22
    lbltiv = icomp
End Sub

Private Sub btncomp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        btncomp_Click (Index)
    ElseIf Button = 2 And btncomp(Index).Tag <> "sxmac" Then
        If btncomp(Index).Tag <> "no" Then
            If icomp > 0 Then
                btncomp(Index).Picture = imglst1.ListImages(2).Picture
                btncomp(Index).Tag = "no"
                icomp = icomp - 1
            End If
        ElseIf btncomp(Index).Tag <> "" Then
            If icomp < 22 Then
                btncomp(Index).Picture = LoadPicture()
                btncomp(Index).Tag = ""
                icomp = icomp + 1
                
            End If
        End If
    End If
    lbltiv = icomp
End Sub

Private Sub btnexit_Click()
    If MsgBox("Do you want to exit?", vbYesNo, "Minesweeper") = vbYes Then
        End
    End If
End Sub

Private Sub btnstart_Click()
    lblardyunq = ""
    If btnstart.Caption = "Start" Then
        If cmbmin < 1 Then
            MsgBox "Please input 0 > number", vbCritical, "Minesweeper"
        ElseIf cmbmin > 60 Then
            MsgBox "Please input 60 < number", vbCritical, "Minesweeper"
        Else
        frtiv.Visible = True
        img1.Visible = True
        tmrtime.Enabled = True
        lbltime.Visible = True
        frmin.Visible = False
        v = 60
        r = cmbmin - 1
        If cmbmin < 10 Then s = 0
            lbltime = s & cmbmin & " : " & "00"
            piccomp.Scale (1, 1)-(11, 11)
            iindex = 0
            For i = 1 To 10
                For j = 1 To 10
                    Load btncomp(iindex)
                    btncomp(iindex).Left = j
                    btncomp(iindex).Top = i
                    btncomp(iindex).Visible = True
                    X(i, j) = 0
                    iindex = iindex + 1
                Next
            Next
            j = 0
            For i = 0 To 11
               cm(i, j) = -7
            Next
            i = 0
            For j = 0 To 11
                cm(i, j) = -7
            Next
            i = 11
            For j = 0 To 11
                cm(i, j) = -7
            Next
            j = 11
            For i = 0 To 11
                cm(i, j) = -7
            Next
            btnstart.Caption = "Try again"
            comp
        End If
    Else
        frmin.Visible = True
        Erase X()
        piccomp.Enabled = True
        For i = 0 To 99
            Unload btncomp(i)
            tmrtime.Enabled = False
            lbltime.Visible = False
        Next
        icomp = 22
        ii = 0
        o = 0
        frtiv.Visible = False
        btnstart.Caption = "Start"
    End If
End Sub
Public Function comptest() As Boolean
    Dim Index As Integer
    comptest = False
    Randomize Timer
    Index = Rnd * 99
    t = True
    i = Index \ 10 + 1
    j = Index Mod 10 + 1
    If X(i, j) <> -1 Then
        If j + 1 <= 10 Then
            If X(i, j + 1) <> -1 Then
                X(i, j + 1) = X(i, j + 1) + 1
            End If
        End If
        If j - 1 >= 1 Then
            If X(i, j - 1) <> -1 Then
                X(i, j - 1) = X(i, j - 1) + 1
            End If
        End If
        If i - 1 >= 1 Then
            If X(i - 1, j) <> -1 Then
                X(i - 1, j) = X(i - 1, j) + 1
            End If
        End If
        If i + 1 <= 10 Then
            If X(i + 1, j) <> -1 Then
                X(i + 1, j) = X(i + 1, j) + 1
            End If
        End If
        If i - 1 >= 1 And j - 1 >= 1 Then
            If X(i - 1, j - 1) <> -1 Then
                X(i - 1, j - 1) = X(i - 1, j - 1) + 1
            End If
        End If
        If i + 1 <= 10 And j + 1 <= 10 Then
           If X(i + 1, j + 1) <> -1 Then
                X(i + 1, j + 1) = X(i + 1, j + 1) + 1
            End If
        End If
        If i - 1 >= 1 And j + 1 <= 10 Then
            If X(i - 1, j + 1) <> -1 Then
                X(i - 1, j + 1) = X(i - 1, j + 1) + 1
            End If
        End If
        If i + 1 <= 10 And j - 1 >= 1 Then
            If X(i + 1, j - 1) <> -1 Then
                X(i + 1, j - 1) = X(i + 1, j - 1) + 1
            End If
        End If
        btncomp(Index).Caption = ""
        X(i, j) = -1
        comptest = True
        icomp = icomp - 1
    End If
End Function
Private Sub btncomp_Click(Index As Integer)
    i = Index \ 10 + 1
    j = Index Mod 10 + 1
    If btncomp(Index).Tag = "" Then
        If X(i, j) = -1 Then
            btncomp(Index).Picture = imglst1.ListImages(1).Picture
            Lose
        ElseIf X(i, j) = 0 Then
            zs = Index
            zro (zs)
        Else
            btncomp(Index).Caption = X(i, j)
            If X(i, j) > 3 Then
                btncomp(Index).BackColor = vbRed
            ElseIf X(i, j) = 1 Then
                btncomp(Index).BackColor = vbGreen
            ElseIf X(i, j) = 2 Then
                btncomp(Index).BackColor = vbYellow
            ElseIf X(i, j) = 3 Then
               btncomp(Index).BackColor = &H80FF&
            End If
            o = o + 1
            txtb = o
        End If
        btncomp(Index).Tag = "sxmac"
    End If
    txtb = o
    If o + icomp = 100 Or 100 - o = 22 Then
        win
    End If
End Sub

Public Sub zro(zs)
    i = zs \ 10 + 1
    j = zs Mod 10 + 1
    If btncomp(zs).Tag <> "sxmac" Then
        test (zs)
    End If
        bindex = (i - 1) * 10 + j + 1 - 1
    If cm(i, j + 1) <> -7 Then
        If btncomp(bindex).Tag = "" Then
            If btncomp(bindex).Tag <> "sxmac" Then
                btncomp(bindex).Tag = "z"
            End If
            test (bindex)
            If X(i, j + 1) = 0 Then
                zro (bindex)
            End If
        End If
    End If
    i = zs \ 10 + 1
    j = zs Mod 10 + 1
   bindex = (i - 1) * 10 + j - 1 - 1
   If cm(i, j - 1) <> -7 Then
        If btncomp(bindex).Tag = "" Then
            If btncomp(bindex).Tag <> "sxmac" Then
                btncomp(bindex).Tag = "z"
            End If
                test (bindex)
            If X(i, j - 1) = 0 Then
                zro (bindex)
            End If
        End If
    End If
    i = zs \ 10 + 1
    j = zs Mod 10 + 1
    bindex = (i - 1 - 1) * 10 + j - 1
    If cm(i - 1, j) <> -7 Then
        If btncomp(bindex).Tag = "" Then
            If btncomp(bindex).Tag <> "sxmac" Then
                btncomp(bindex).Tag = "z"
            End If
                test (bindex)
            If X(i - 1, j) = 0 Then
                zro (bindex)
            End If
        End If
    End If
    bindex = (i - 1 + 1) * 10 + j - 1
    If cm(i + 1, j) <> -7 Then
        If btncomp(bindex).Tag = "" Then
            If btncomp(bindex).Tag <> "sxmac" Then
                btncomp(bindex).Tag = "z"
            End If
                test (bindex)
            If X(i + 1, j) = 0 Then
                zro (bindex)
            End If
        End If
    End If
End Sub



Public Sub test(ind)
    If ind <> 100 Then
        ii = ind \ 10 + 1
        jj = ind Mod 10 + 1
    Else
        ii = ind \ 10
        jj = 10
    End If
    btncomp(ind).Caption = X(ii, jj)
    If X(ii, jj) > 3 Then
        btncomp(ind).BackColor = vbRed
    ElseIf X(ii, jj) = 1 Then
        btncomp(ind).BackColor = vbGreen
    ElseIf X(ii, jj) = 2 Then
        btncomp(ind).BackColor = vbYellow
    ElseIf X(ii, jj) = 3 Then
       btncomp(ind).BackColor = &H80FF&
    ElseIf X(ii, jj) = 0 Then
        btncomp(ind).BackColor = vbWhite
    ElseIf X(ii, jj) = -1 Then
        btncomp(ind).BackColor = vbBlack
    End If
    btncomp(ind).Tag = "sxmac"
    o = o + 1
    
End Sub

Private Sub Form_Load()
    cmbmin.AddItem 1
    cmbmin.AddItem 2
    cmbmin.AddItem 3
    cmbmin.AddItem 5
    cmbmin.AddItem 10
    cmbmin.AddItem 15
    lblcoopyright = Chr(169) & " Melqon Hovhannisyan"
End Sub



Private Sub tmrtime_Timer()
    v = v - 1
    If v = 0 Then
        If r > 0 Then
            r = r - 1
            v = 60
        Else
            Lose
        End If
    End If
    If v < 10 Then s = 0
    If r < 10 Then sr = 0
    lbltime = sr & r & " : " & s & v
End Sub


Public Sub Lose()
    piccomp.Enabled = False
    iindex = 0
    For i = 1 To 10
        For j = 1 To 10
        If btncomp(iindex).Tag = "no" Then
            btncomp(iindex).Picture = LoadPicture()
        End If
        If X(i, j) = -1 Then
            btncomp(iindex).Picture = imglst1.ListImages(1).Picture
        ElseIf X(i, j) > 3 Then
            btncomp(iindex).BackColor = vbRed
        ElseIf X(i, j) = 1 Then
            btncomp(iindex).BackColor = vbGreen
        ElseIf X(i, j) = 2 Then
            btncomp(iindex).BackColor = vbYellow
        ElseIf X(i, j) = 3 Then
           btncomp(iindex).BackColor = &H80FF&
        ElseIf X(i, j) = 0 Then
            btncomp(iindex).BackColor = vbWhite
        End If
        If X(i, j) <> -1 Then
            btncomp(iindex).Caption = X(i, j)
        End If
        iindex = iindex + 1
    Next
Next
    btnstart.Visible = True
    tmrtime.Enabled = False
    lblardyunq = "You lose..."
End Sub

Public Sub win()
    iindex = 0
    piccomp.Enabled = False
    For i = 1 To 10
        For j = 1 To 10
            If btncomp(iindex).Tag = "" Then
                btncomp(iindex).Picture = imglst1.ListImages(2).Picture
                iindex = iindex + 1
            End If
        Next
    Next
    btnstart.Visible = True
    tmrtime.Enabled = False
    lblardyunq = "You Win!!"
End Sub
