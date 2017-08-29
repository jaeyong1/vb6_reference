VERSION 5.00
Begin VB.Form frm출력 
   Caption         =   "출력화면"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmd종료 
      Caption         =   "종    료"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   3720
      Width           =   2295
   End
End
Attribute VB_Name = "frm출력"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd종료_Click()
End
End Sub

Private Sub Form_Activate()
Dim i%
Dim 합계&, 평균!
Dim a!, b!, c!

frm출력.Cls
frm출력.Print ""
frm출력.Print "   번  호    이  름    어셈블리    코 볼    포트란    총 점    평 균"
frm출력.Print "----------------------------------------------------"

For i = 1 To cnt
frm출력.Print Tab(3); sung(i, 1);
frm출력.Print Tab(9); sung(i, 2);
frm출력.Print Tab(18); sung(i, 3);
a = a + sung(i, 3)
frm출력.Print Tab(27); sung(i, 4);
b = b + sung(i, 4)
frm출력.Print Tab(34); sung(i, 5);
c = c + sung(i, 5)
합계 = sung(i, 3) + sung(i, 4) + sung(i, 5)
frm출력.Print Tab(41); 합계;
평균 = 합계 / 3
frm출력.Print Tab(48); Format(평균, "##0.00");
Next i
frm출력.Print ""
frm출력.Print "----------------------------------------------------"

frm출력.Print Tab(18); Format(a / cnt, "#0.00");
frm출력.Print Tab(27); Format(b / cnt, "#0.00");
frm출력.Print Tab(34); Format(c / cnt, "#0.00");

End Sub
