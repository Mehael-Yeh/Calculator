VERSION 5.00
Begin VB.Form Yale的初代计算器 
   Caption         =   "计算器  v0.4  Bulid 07/06/2020"
   ClientHeight    =   4560
   ClientLeft      =   7635
   ClientTop       =   3300
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   4755
   Begin VB.CommandButton Cmd2 
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   1750
      TabIndex        =   20
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   1080
      TabIndex        =   19
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "√"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   360
      TabIndex        =   18
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Cmd4 
      Caption         =   "←"
      Height          =   615
      Left            =   3840
      TabIndex        =   17
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "^3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   3840
      TabIndex        =   16
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "ln"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   3120
      TabIndex        =   15
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "log"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   2400
      TabIndex        =   14
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "e^x"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   3120
      TabIndex        =   13
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   2400
      TabIndex        =   12
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "|x|"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   2400
      TabIndex        =   11
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "10^x"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3120
      TabIndex        =   10
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "^x"
      Height          =   495
      Index           =   6
      Left            =   3840
      TabIndex        =   9
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "^2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3840
      TabIndex        =   8
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "Mod"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2400
      TabIndex        =   7
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3840
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Cmd3 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Yale的初代计算器"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X1, X2 As Double: Public Y As String: Public a, p As Boolean
Private Sub Cmd1_Click(Index As Integer)
    If Index <> 12 Then Text1.Text = Text1.Text & Cmd1(Index).Caption
    If Index = 11 Then
        Cmd1(11).Enabled = False
    End If
    If Index = 12 Then
        X2 = Val(Text1.Text)
        If Y = "+" Then X1 = X1 + X2
        If Y = "-" Then X1 = X1 - X2
        If Y = "×" Then X1 = X1 * X2
        If Y = "÷" Then X1 = X1 / X2
        If Y = "Mod" Then X1 = X1 Mod X2
        If Y = "^2" Then X1 = X1 * X1
        If Y = "^3" Then X1 = X1 ^ 3
        If Y = "^x" Then X1 = X1 ^ X2
        If Y = "10^x" Then X1 = 10 ^ X1
        If Y = "^x" Then X1 = Abs(X1)
        If Y = "e^x" Then X1 = Exp(X1)
        If Y = "log" Then
            If X1 <= 0 Then
                MsgBox "无效数值", 0, "警告"
                Text1.Text = X1
                Text1.SetFocus
            Else
                X1 = Log(X1)
            End If
        End If
        If Y = "ln" Then
             If X1 <= 0 Then
                MsgBox "无效数值", 0, "警告"
                Text1.Text = X1
                Text1.SetFocus
            Else
                X1 = Log(X1) / Log(2.718281828459)
            End If
        End If
        If Y = "√" Then
            If X1 < 0 Then
                MsgBox "无效数值", 0, "警告"
                Text1.Text = X1
                Text1.SetFocus
            Else
                X1 = Sqr(X1)
            End If
        End If
        If Y = "Sin" Then X1 = Sin(X1)
        If Y = "Cos" Then X1 = (X1)
        Text1.Text = X1
        If InStr(Text1.Text, ".") = 0 Then
            Cmd1(11).Enabled = True
        Else
            Cmd1(11).Enabled = False
        End If
    End If
End Sub

Private Sub Cmd2_Click(Index As Integer)
    Cmd1(11).Enabled = True
    If Index = 9 Then
        If Sgn(Val(Text1.Text)) = -1 Then
            Text1.Text = -Val(Text1.Text)
        ElseIf Sgn(Val(Text1.Text)) = 1 Then
            Text1.Text = "-" & Text1.Text
        ElseIf Sgn(Val(Text1.Text)) = 0 Then
            If a = False Then
                Text1.Text = "-"
                a = True
            Else
                Text1.Text = ""
                a = False
            End If
        End If
    Else
        Y = Cmd2(Index).Caption
        X1 = Val(Text1.Text)
        Text1.Text = ""
    End If
End Sub

Private Sub Cmd3_Click()
    Text1.Text = ""
    Cmd1(11).Enabled = True
End Sub

Private Sub Cmd4_Click()
     If Text1.Text <> "" Then
        Last = Right(Text1.Text, 1)
        Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
        If Last = "." Then
            Cmd1(11).Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    For k = 1 To 12
        Load Cmd1(k)
    Next k
    mtop = 1080
    mwidth = 500
    mheight = 500
    For i = 1 To 4
        mleft = 360
        For j = 1 To 3
            k = (i - 1) * 3 + j
            Cmd1(k).Caption = k
            If k = 10 Then k = 0: Cmd1(k).Caption = k
            If k = 11 Then Cmd1(k).Caption = "."
            If k = 12 Then Cmd1(k).Caption = "="
            Cmd1(k).Visible = True
            Cmd1(k).Width = mwidth
            Cmd1(k).Height = mheight
            Cmd1(k).Left = mleft
            Cmd1(k).Top = mtop
            Cmd1(k).FontSize = 12
            Cmd1(k).FontBold = True
            mleft = mleft + mwidth + 200
        Next j
        mtop = mtop + mheight + 220
    Next i
    MsgBox "版本v0.4" & vbCrLf & "修复内容：" & vbCrLf & "小数点输入修复" & vbCrLf & vbCrLf & "本软件仅供学习使用，有任何意见可联系作者QQ:1175592624", 0, "使用提示"
End Sub
