VERSION 5.00
Begin VB.Form frmMIDI 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "날달걀 :)"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   1530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdStop 
      Caption         =   "■"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "||"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "▶"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "frmMIDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MDPlayer As IMediaControl, MDPosition As IMediaPosition
Private Sub cmdPause_Click()
MDPlayer.Pause
cmdPlay.Enabled = True
cmdPause.Enabled = False
End Sub

Private Sub cmdPlay_Click()
MDPlayer.Run
cmdPlay.Enabled = 0
cmdPause.Enabled = 1
End Sub

Private Sub cmdStop_Click()
MDPlayer.Stop
MDPosition.CurrentPosition = 0
cmdPlay.Enabled = 1
cmdPause.Enabled = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Dim MIDIPath As String
    
    '
    ' 미디 파일 경로를 만듭니다.
    '
    MIDIPath = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") & "Beethoven_Virus.mid"
    
    '
    ' quartz.dll 재생기를 로드합니다.
    '
    Set MDPlayer = New FilgraphManager
    
    '
    ' 파일을 읽습니다.
    '
    MDPlayer.RenderFile MIDIPath
    
    '
    ' 위치 조절을 위한 컨트롤을 캐스팅합니다.
    '
    Set MDPosition = MDPlayer
    
    '
    ' MIDI 재생
    '
    MDPosition.CurrentPosition = 0
    MDPlayer.Run
    cmdPlay.Enabled = False
If Not Err.Number = 0 Then MsgBox Err.Number & Err.Description
End Sub

Private Sub Image1_DblClick()
On Error Resume Next
MDPlayer.Stop
MDPosition.CurrentPosition = 0
cmdPlay.Enabled = 1
cmdPause.Enabled = 0
'이스터 에그 안 이스터 에그 :)
MIDIPath = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") & "NF.mid"
MDPlayer.RenderFile MIDIPath
MDPosition.CurrentPosition = 0
MDPlayer.Run
cmdPlay.Enabled = 0
cmdPause.Enabled = 1
End Sub
