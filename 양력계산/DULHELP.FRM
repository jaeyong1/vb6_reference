VERSION 5.00
Begin VB.Form dulhelp 
   ClientHeight    =   4575
   ClientLeft      =   2115
   ClientTop       =   1980
   ClientWidth     =   6105
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4575
   ScaleWidth      =   6105
   Begin VB.CommandButton Command1 
      Caption         =   "다 읽었습니다"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "<사용 방법>"
      BeginProperty Font 
         Name            =   "바탕"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "바탕"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "dulhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Label2 = "제작자의 실력한계로 자동으로 데이터를 추출하지 못하는데 정말로 죄송합니다. 파일불러오기를 눌러서 데이터 파일의 내용을 보면서 입력할수 있게 했습니다.불러오는 파일은 C:\airfoil.dat로 고정되어 있습니다. 많은 양해 부탁드립니다 ^.^ " + Chr(13) + Chr(13) + "파일불러오기를 했을때 줄수가 1개인 경우가 있는데, 이때는 edit실행 - [file]-[open]-[c:\airfoil.dat]-[save]-[exit]를 해주시면 됩니다."
End Sub

