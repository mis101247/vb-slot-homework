VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "拉霸機"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   9450
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000012&
      Caption         =   "一萬次"
      Height          =   495
      Left            =   3718
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   15
      Top             =   6315
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "清除歷史"
      Height          =   495
      Left            =   6720
      TabIndex        =   10
      Top             =   3705
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   364
      Left            =   7140
      TabIndex        =   9
      Text            =   "0"
      Top             =   5190
      Width           =   1066
   End
   Begin VB.CommandButton Command4 
      Caption         =   "按開始後在按我改變"
      Height          =   715
      Left            =   8190
      TabIndex        =   8
      Top             =   4830
      Width           =   1066
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   6720
      TabIndex        =   7
      Top             =   1287
      Width           =   2262
   End
   Begin VB.TextBox Text1 
      Height          =   364
      Left            =   7140
      TabIndex        =   6
      Text            =   "0"
      Top             =   4830
      Width           =   1066
   End
   Begin VB.Timer Timer1 
      Left            =   3060
      Top             =   2625
   End
   Begin VB.Timer Timer2 
      Left            =   4110
      Top             =   2625
   End
   Begin VB.Timer Timer3 
      Left            =   5280
      Top             =   2625
   End
   Begin VB.TextBox money2 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   481
      Left            =   2760
      TabIndex        =   5
      Text            =   "0"
      Top             =   4920
      Width           =   3172
   End
   Begin VB.CommandButton Command3 
      Caption         =   "檢視次數"
      Height          =   715
      Left            =   7140
      TabIndex        =   4
      Top             =   5535
      Width           =   1183
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   5382
      Top             =   6318
   End
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3718
      TabIndex        =   3
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000012&
      Caption         =   "開始"
      Height          =   495
      Left            =   3718
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox money1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   442
      Left            =   2760
      TabIndex        =   1
      Text            =   "50000"
      Top             =   1053
      Width           =   3237
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "出現/得彩金"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   611
      Left            =   6669
      TabIndex        =   16
      Top             =   819
      Width           =   2353
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "下注金額"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   27.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2760
      TabIndex        =   14
      Top             =   4320
      Width           =   3195
   End
   Begin VB.Label label4 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "彩金金額"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   27.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   611
      Left            =   2691
      TabIndex        =   13
      Top             =   468
      Width           =   3198
   End
   Begin VB.Label Label3 
      Caption         =   "後門區(測試)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   25.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6315
      TabIndex        =   12
      Top             =   4245
      Width           =   2820
   End
   Begin VB.Label Label2 
      Caption         =   "出現次數:                      數字:"
      Height          =   720
      Left            =   6315
      TabIndex        =   11
      Top             =   4830
      Width           =   825
   End
   Begin VB.Image RP 
      Height          =   2805
      Index           =   0
      Left            =   2760
      Picture         =   "拉霸測試.frx":0000
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Image RP 
      Height          =   2805
      Index           =   1
      Left            =   3870
      Picture         =   "拉霸測試.frx":2C98
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Image RP 
      Height          =   2805
      Index           =   2
      Left            =   4980
      Picture         =   "拉霸測試.frx":5930
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   6600
      Left            =   120
      Picture         =   "拉霸測試.frx":85C8
      Top             =   465
      Width           =   2475
   End
   Begin VB.Image P 
      Height          =   2805
      Index           =   9
      Left            =   -465
      Picture         =   "拉霸測試.frx":9892
      Top             =   6435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image P 
      Height          =   2805
      Index           =   8
      Left            =   -465
      Picture         =   "拉霸測試.frx":C8DA
      Top             =   6435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image P 
      Height          =   2805
      Index           =   7
      Left            =   -465
      Picture         =   "拉霸測試.frx":F2BC
      Top             =   6435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image P 
      Height          =   2805
      Index           =   6
      Left            =   -465
      Picture         =   "拉霸測試.frx":11CC8
      Top             =   6435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image P 
      Height          =   2805
      Index           =   5
      Left            =   -465
      Picture         =   "拉霸測試.frx":14A4C
      Top             =   6435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image P 
      Height          =   2805
      Index           =   4
      Left            =   -465
      Picture         =   "拉霸測試.frx":176D5
      Top             =   6435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image P 
      Height          =   2805
      Index           =   3
      Left            =   -465
      Picture         =   "拉霸測試.frx":1A041
      Top             =   6435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image P 
      Height          =   2805
      Index           =   2
      Left            =   -465
      Picture         =   "拉霸測試.frx":1CB54
      Top             =   6435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image P 
      Height          =   2805
      Index           =   1
      Left            =   -465
      Picture         =   "拉霸測試.frx":1F18C
      Top             =   6435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image P 
      Height          =   2805
      Index           =   0
      Left            =   -465
      Picture         =   "拉霸測試.frx":2178D
      Top             =   6435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "說明:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   25.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   117
      TabIndex        =   0
      Top             =   0
      Width           =   1287
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim money As Double '錢盒
Dim Win(9) As Integer '出現數字的次數
Dim NTimer As Integer 'timer1~3跑完的+1

Private Sub Command1_Click()
If money >= money2 And money2 >= 0 Then
money = money - money2
Timer1.Interval = 1
Timer2.Interval = 1
Timer3.Interval = 1
NTimer = 0
Timer4.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
For i = 0 To 9
Win(i) = 0
Next i
Else
MsgBox "彩金不足OR下注金額錯誤", , "彩金不足OR下注金額錯誤"
End If
End Sub

Private Sub Command2_Click()
Timer1.Interval = 101
Timer2.Interval = 101
Timer3.Interval = 101
End Sub

Private Sub Command3_Click() '測試項
For i = 0 To 9
Print Win(i);
Next i
Print NTimer
End Sub

Private Sub Command4_Click() '測試項
Dim i As Integer
i = Text2.Text
Win(i) = Text1.Text
End Sub

Private Sub Command5_Click()
List1.Clear
End Sub

Private Sub Command6_Click()
If money >= money2 And money2 >= 0 Then
Randomize
money = money - money2
For W = 1 To 10000
x = Int(Rnd * 10)
y = Int(Rnd * 10)
z = Int(Rnd * 10)
RP(0).Picture = P(x).Picture
RP(1).Picture = P(y).Picture
RP(2).Picture = P(z).Picture
For j = 0 To 9
  Win(j) = 0
  If RP(0).Picture = P(j).Picture Then Win(j) = Win(j) + 1
  If RP(1).Picture = P(j).Picture Then Win(j) = Win(j) + 1
  If RP(2).Picture = P(j).Picture Then Win(j) = Win(j) + 1
Next j
  For i = 0 To 9
   If Win(i) = 1 And i = 7 Then
    money = money + (money2 * 1.5)
    List1.AddItem i & "      " & money - money1
    money1 = money
   Else
     If Win(i) = 2 Then
     If i = 7 Then money = money + (money2 * 50)
     If i = 1 Or i = 2 Or i = 3 Or i = 4 Or i = 5 Or i = 6 Or i = 8 Or i = 9 Then money = money + (money2 * 1.5)
    List1.AddItem i & i & "    " & money - money1
     money1 = money
    Else
     If Win(i) = 3 Then
      If i = 7 Then money = money + (money2 * 100)
      If i = 8 Then money = money + (money2 * 8)
      If i = 6 Then money = money + (money2 * 6)
      If i = 0 Then money = money + (money2 * 0.5)
      If i = 1 Or i = 2 Or i = 3 Or i = 4 Or i = 5 Or i = 9 Then money = money + (money2 * 4)
      List1.AddItem i & i & i & "  " & money - money1
      money1 = money
     End If
    End If
   End If
 Next i
Next W
Else
MsgBox "彩金不足OR下注金額錯誤", , "彩金不足OR下注金額錯誤"
End If
End Sub

Private Sub Form_Load()
money = 50000
End Sub

Private Sub Timer1_Timer()
Randomize
x = Int(Rnd * 10)
RP(0).Picture = P(x).Picture
Timer1.Interval = Timer1.Interval + x
If Timer1.Interval > 100 Then
 Timer1.Interval = 0
 NTimer = NTimer + 1
  For i = 0 To 9
   If RP(0).Picture = P(i).Picture Then
   Win(i) = Win(i) + 1
   End If
  Next i
End If
End Sub

Private Sub Timer2_Timer()
Randomize
x = Int(Rnd * 10)
RP(1).Picture = P(x).Picture
Timer2.Interval = Timer2.Interval + x
If Timer2.Interval > 100 Then
Timer2.Interval = 0
NTimer = NTimer + 1
 For i = 0 To 9
  If RP(1).Picture = P(i).Picture Then
  Win(i) = Win(i) + 1
  End If
 Next i
End If
End Sub

Private Sub Timer3_Timer()
Randomize
x = Int(Rnd * 10)
RP(2).Picture = P(x).Picture
Timer3.Interval = Timer3.Interval + x
If Timer3.Interval > 100 Then
 Timer3.Interval = 0
  NTimer = NTimer + 1
  For i = 0 To 9
   If RP(2).Picture = P(i).Picture Then
   Win(i) = Win(i) + 1
   End If
  Next i
End If
End Sub

Private Sub Timer4_Timer()
money1 = money

If NTimer = 3 Then
Command1.Enabled = True
Command2.Enabled = False
 For i = 0 To 9
  If Win(i) = 1 And i = 7 Then
   money = money + (money2 * 1.5)
   List1.AddItem i & "      " & money - money1
   money1 = money
   Timer4.Enabled = False
  Else
   If Win(i) = 2 Then
    If i = 7 Then money = money + (money2 * 50)
    If i = 1 Or i = 2 Or i = 3 Or i = 4 Or i = 5 Or i = 6 Or i = 8 Or i = 9 Then money = money + (money2 * 1.5)
   List1.AddItem i & i & "    " & money - money1
    money1 = money
    Timer4.Enabled = False
   Else
    If Win(i) = 3 Then
     If i = 7 Then money = money + (money2 * 100)
     If i = 8 Then money = money + (money2 * 8)
     If i = 6 Then money = money + (money2 * 6)
     If i = 0 Then money = money + (money2 * 0.5)
     If i = 1 Or i = 2 Or i = 3 Or i = 4 Or i = 5 Or i = 9 Then money = money + (money2 * 4)
   List1.AddItem i & i & i & "  " & money - money1
     money1 = money
     Timer4.Enabled = False
    End If
   End If
  End If
 Next i
End If
End Sub

