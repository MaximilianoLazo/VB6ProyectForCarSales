VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Cargando 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11880
      Top             =   1320
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7680
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3480
      TabIndex        =   2
      Top             =   8160
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CARGANDO %"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      TabIndex        =   1
      Top             =   8160
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   8775
      Left            =   0
      Picture         =   "Cargando.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "Cargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Rem redimensionar imagen

Image1.Width = Cargando.Width
Image1.Height = Cargando.Height

End Sub
Private Sub Form_Load()
Timer1.Enabled = True

End Sub



Private Sub Timer1_Timer()
Timer1.Interval = Rnd * 300 + 10
ProgressBar1.Value = ProgressBar1.Value + 20
Label2.Caption = ProgressBar1.Value
If Label2.Caption = 100 Then
Unload Me
MenuPrincipal.Show

End If

End Sub

