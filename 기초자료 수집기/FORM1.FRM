VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "안녕하세요.."
   ClientHeight    =   5955
   ClientLeft      =   420
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton gather 
      Caption         =   "수 집 하 기(기본)"
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton gather1 
      Caption         =   "수 집 하 기(상세)"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   7935
      Begin VB.TextBox txt 
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "입니다."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5760
         TabIndex        =   9
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "저는"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   8
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "***** 안녕하세요.. *****"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   3900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "실행에 앞서서 정체를 밝혀주세요 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   720
         TabIndex        =   5
         Top             =   1440
         Width           =   3525
      End
   End
   Begin VB.CommandButton cmdend 
      Cancel          =   -1  'True
      Caption         =   "종 료 하 기"
      Height          =   615
      Left            =   7080
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton virus 
      Caption         =   "바이러스검사"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdhlp 
      Caption         =   "이거는..."
      Default         =   -1  'True
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Made in Korea 재용 엔터테인먼트"
      Height          =   180
      Left            =   6480
      TabIndex        =   1
      Top             =   0
      Width           =   2805
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdend_Click()
MsgBox "쓰기방지를 해주세요 (디스켓의 구멍열기)"
End
End Sub

Private Sub cmdhlp_Click()
Dialog.Show
End Sub
Private Sub virus_Click()
retval = Shell("a:\v3 c:")
End Sub

Private Sub gather1_Click()
MsgBox "쫌 오래걸릴수도 있어요~"
retval = Shell("a:\gather1.bat")
End Sub

Private Sub gather_Click()
FileName = "a:\사용자.txt" '기록파일이름설정
filenum = FreeFile
Open FileName For Append As FreeFile 'append모드 : 있으면 뒤에써내려간다.
Print #filenum, txt; '텍스트박스안의내용기록후 줄바꿈안함 => ; 붙임
Print #filenum, Date; Time; '날짜시간기록
Close #filenum '파일닫기
retval = Shell("a:\gather.bat")
End Sub

