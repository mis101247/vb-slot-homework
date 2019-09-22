VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "佑佑拉霸機"
   ClientHeight    =   5993
   ClientLeft      =   65
   ClientTop       =   455
   ClientWidth     =   8905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5993
   ScaleWidth      =   8905
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "認輸~從玩!"
      Height          =   364
      Left            =   7956
      TabIndex        =   11
      Top             =   5616
      Width           =   949
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000012&
      Caption         =   "一萬次(無扣錢)"
      Height          =   495
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "清除歷史"
      Height          =   495
      Left            =   6240
      TabIndex        =   6
      Top             =   3345
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2249
      Left            =   6240
      TabIndex        =   5
      Top             =   930
      Width           =   2262
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
         Size            =   18.34
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   481
      Left            =   2760
      TabIndex        =   4
      Text            =   "0"
      Top             =   4920
      Width           =   3172
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   0
      Top             =   5616
   End
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000012&
      Caption         =   "開始"
      Height          =   495
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox money1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   19.7
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
         Size            =   18.34
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6195
      TabIndex        =   10
      Top             =   465
      Width           =   2355
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "下注金額"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   27.85
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2760
      TabIndex        =   8
      Top             =   4320
      Width           =   3195
   End
   Begin VB.Label label4 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "彩金金額"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   27.85
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   611
      Left            =   2691
      TabIndex        =   7
      Top             =   468
      Width           =   3198
   End
   Begin VB.Image RP 
      Height          =   2431
      Index           =   0
      Left            =   2756
      Picture         =   "拉霸.frx":0000
      Top             =   1586
      Width           =   975
   End
   Begin VB.Image RP 
      Height          =   2431
      Index           =   1
      Left            =   3874
      Picture         =   "拉霸.frx":2C98
      Top             =   1586
      Width           =   975
   End
   Begin VB.Image RP 
      Height          =   2431
      Index           =   2
      Left            =   4979
      Picture         =   "拉霸.frx":5930
      Top             =   1586
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   5720
      Left            =   117
      Picture         =   "拉霸.frx":85C8
      Top             =   468
      Width           =   2145
   End
   Begin VB.Image P 
      Height          =   2431
      Index           =   9
      Left            =   416
      Picture         =   "拉霸.frx":9829
      Top             =   5473
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image P 
      Height          =   2431
      Index           =   8
      Left            =   416
      Picture         =   "拉霸.frx":C871
      Top             =   5473
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image P 
      Height          =   2431
      Index           =   7
      Left            =   416
      Picture         =   "拉霸.frx":F253
      Top             =   5473
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image P 
      Height          =   2431
      Index           =   6
      Left            =   416
      Picture         =   "拉霸.frx":11C5F
      Top             =   5473
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image P 
      Height          =   2431
      Index           =   5
      Left            =   416
      Picture         =   "拉霸.frx":149E3
      Top             =   5473
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image P 
      Height          =   2431
      Index           =   4
      Left            =   416
      Picture         =   "拉霸.frx":1766C
      Top             =   5473
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image P 
      Height          =   2431
      Index           =   3
      Left            =   416
      Picture         =   "拉霸.frx":19FD8
      Top             =   5473
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image P 
      Height          =   2431
      Index           =   2
      Left            =   416
      Picture         =   "拉霸.frx":1CAEB
      Top             =   5473
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image P 
      Height          =   2431
      Index           =   1
      Left            =   416
      Picture         =   "拉霸.frx":1F123
      Top             =   5473
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image P 
      Height          =   2431
      Index           =   0
      Left            =   416
      Picture         =   "拉霸.frx":21724
      Top             =   5473
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "說明:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   25.81
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   234
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

Private Sub Command3_Click()
money = 50000
money1 = 50000
List1.Clear
End Sub

Private Sub Command5_Click()
List1.Clear
End Sub

Private Sub Command6_Click()
Randomize
If money >= money2 And money2 >= 0 Then
For yo = 1 To 10000
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
     If i = 0 Then money = money * 1 + money2 * 1
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
Next yo
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

