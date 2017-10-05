VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SimTraffic Config"
   ClientHeight    =   1860
   ClientLeft      =   3810
   ClientTop       =   6525
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Contador 
      Enabled         =   0   'False
      Left            =   3075
      Top             =   780
   End
   Begin VB.CommandButton cmdContinuar 
      Caption         =   "Continuar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1995
      TabIndex        =   9
      Top             =   1335
      Width           =   1185
   End
   Begin VB.CommandButton cmdPausar 
      Caption         =   "Pausar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3210
      TabIndex        =   6
      Top             =   1335
      Width           =   1185
   End
   Begin VB.CommandButton cmdParar 
      Caption         =   "Parar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1335
      Width           =   1185
   End
   Begin VB.TextBox txtTempo 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2430
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "5"
      Top             =   825
      Width           =   375
   End
   Begin VB.CommandButton cmdIniciar 
      Caption         =   "Iniciar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   735
      TabIndex        =   1
      Top             =   1335
      Width           =   1185
   End
   Begin MSComctlLib.Slider quantCarros 
      Height          =   270
      Left            =   30
      TabIndex        =   0
      Top             =   405
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   476
      _Version        =   393216
      Max             =   20
      SelStart        =   4
      Value           =   4
   End
   Begin MSComctlLib.Slider tempoSemaforo 
      Height          =   285
      Left            =   2295
      TabIndex        =   7
      Top             =   405
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   503
      _Version        =   393216
      LargeChange     =   500
      SmallChange     =   500
      Min             =   1000
      Max             =   6000
      SelStart        =   3000
      TickFrequency   =   500
      Value           =   3000
   End
   Begin VB.Label lblSemaforo 
      Caption         =   "Velocidade do semáforo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   2385
      TabIndex        =   8
      Top             =   150
      Width           =   2940
   End
   Begin VB.Label Label2 
      Caption         =   "Tempo de Simulação (min):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   915
      Width           =   2355
   End
   Begin VB.Label lblCarros 
      Caption         =   "Número de Carros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   150
      Width           =   2130
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdContinuar_Click()
    frmSimA.continua
    frmSimB.continua
    cmdContinuar.Enabled = False
    cmdPausar.Enabled = True
    Contador.Enabled = True
End Sub

Private Sub cmdIniciar_Click()
    
    frmSimA.Inicializa_Mundo
    
    frmSimB.Inicializa_Mundo
    
    cmdIniciar.Enabled = False
    cmdContinuar.Enabled = False
    cmdPausar.Enabled = True
    cmdParar.Enabled = True
    
    Contador.Interval = 60000
    Contador.Enabled = True
    
    txtTempo.Enabled = False
    
End Sub

Private Sub cmdPausar_Click()
    frmSimA.pausa
    frmSimB.pausa
    cmdContinuar.Enabled = True
    cmdPausar.Enabled = False
    Contador.Enabled = False
End Sub

Private Sub Contador_Timer()
    txtTempo.Text = txtTempo.Text - 1
    If txtTempo.Text = "0" Then MsgBox "Fim"
End Sub

Private Sub Form_Load()
    MAX_CARROS = quantCarros.Value
    lblCarros.Caption = "Número de Carros: " & MAX_CARROS
    
    TEMPO_SEMAFORO = tempoSemaforo.Value
    lblSemaforo.Caption = "Velocidade do semáforo(segs): " & TEMPO_SEMAFORO / 1000
    TEMPO_SINAL = TEMPO_SEMAFORO / 2
    frmSimA.tsemA.Interval = TEMPO_SINAL
    frmSimA.tsemB.Interval = TEMPO_SEMAFORO
    
    frmSimB.tsemA.Interval = TEMPO_SINAL
    frmSimB.tsemB.Interval = TEMPO_SEMAFORO
    
    frmSimA.Show
    frmSimA.SetFocus
    frmSimB.Show
    frmSimB.SetFocus
    
End Sub

Private Sub quantCarros_Change()

    MAX_CARROS = quantCarros.Value
    lblCarros.Caption = "Número de Carros: " & MAX_CARROS

End Sub

Private Sub Text1_Change()

End Sub

Private Sub tempoSemaforo_Change()
    
    TEMPO_SEMAFORO = tempoSemaforo.Value
    lblSemaforo.Caption = "Velocidade do semáforo(segs): " & TEMPO_SEMAFORO / 1000
    TEMPO_SINAL = TEMPO_SEMAFORO / 2
    
    frmSimA.tsemA.Interval = TEMPO_SINAL
    frmSimA.tsemB.Interval = TEMPO_SEMAFORO
    
    frmSimB.tsemA.Interval = TEMPO_SINAL
    frmSimB.tsemB.Interval = TEMPO_SEMAFORO
    
End Sub

Private Sub txtTempo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
       Debug.Print KeyAscii
End Sub
