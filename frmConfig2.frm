VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SimTraffic Config"
   ClientHeight    =   1725
   ClientLeft      =   3795
   ClientTop       =   7455
   ClientWidth     =   5820
   Icon            =   "frmConfig2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5820
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
      Left            =   825
      TabIndex        =   4
      Top             =   1275
      Width           =   1185
   End
   Begin VB.TextBox txtTempo 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2535
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "5"
      Top             =   765
      Width           =   375
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
      Left            =   4545
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1275
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
      Left            =   3315
      TabIndex        =   1
      Top             =   1275
      Width           =   1185
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
      Left            =   2100
      TabIndex        =   0
      Top             =   1275
      Width           =   1185
   End
   Begin VB.Timer Contador 
      Enabled         =   0   'False
      Left            =   3180
      Top             =   720
   End
   Begin MSComctlLib.Slider quantCarros 
      Height          =   270
      Left            =   135
      TabIndex        =   5
      Top             =   345
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
      Left            =   2400
      TabIndex        =   6
      Top             =   345
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
      Left            =   180
      TabIndex        =   9
      Top             =   90
      Width           =   2130
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
      Left            =   195
      TabIndex        =   8
      Top             =   855
      Width           =   2355
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
      Left            =   2490
      TabIndex        =   7
      Top             =   90
      Width           =   2940
   End
End
Attribute VB_Name = "frmConfig2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdContinuar_Click()
    cmdIniciar.Enabled = False
    cmdContinuar.Enabled = False
    cmdPausar.Enabled = True
    cmdParar.Enabled = True
    frmSimC.continue
    frmSimD.continue
End Sub

Private Sub cmdIniciar_Click()
    cmdIniciar.Enabled = False
    cmdContinuar.Enabled = False
    cmdPausar.Enabled = True
    cmdParar.Enabled = True
    
    Contador.Interval = 60000
    Contador.Enabled = True
    
    txtTempo.Enabled = False
    
    frmSimC.start
    frmSimD.start
End Sub

Private Sub cmdParar_Click()
    frmSimC.stops
    frmSimD.stops
    cmdIniciar.Enabled = True
    cmdContinuar.Enabled = False
    cmdPausar.Enabled = False
    cmdParar.Enabled = False
    txtTempo.Enabled = True
    Contador.Enabled = False
End Sub

Private Sub cmdPausar_Click()
    cmdIniciar.Enabled = False
    cmdContinuar.Enabled = True
    cmdPausar.Enabled = False
    cmdParar.Enabled = True
    frmSimC.pause
    frmSimD.pause
End Sub

Private Sub Contador_Timer()
    txtTempo.Text = txtTempo.Text - 1
    If txtTempo.Text = "0" Then
        cmdParar_Click
        MsgBox "Fim da Simulação"
    End If
End Sub

Private Sub Form_Load()
    frmSimC.Show
    frmSimD.Show
    
    MAX_CARROS = quantCarros.Value
    lblCarros.Caption = "Número de Carros: " & MAX_CARROS
    
    TEMPO_SEMAFORO = tempoSemaforo.Value
    lblSemaforo.Caption = "Velocidade do semáforo(segs): " & TEMPO_SEMAFORO / 1000
    TEMPO_SINAL = TEMPO_SEMAFORO / 2
    frmSimC.tsemA.Interval = TEMPO_SINAL
    frmSimC.tsemB.Interval = TEMPO_SEMAFORO
    
    frmSimD.tsemA.Interval = TEMPO_SINAL
    frmSimD.tsemB.Interval = TEMPO_SEMAFORO
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub quantCarros_Click()
    MAX_CARROS = quantCarros.Value
    lblCarros.Caption = "Número de Carros: " & MAX_CARROS
End Sub

Private Sub tempoSemaforo_Click()
    TEMPO_SEMAFORO = tempoSemaforo.Value
    lblSemaforo.Caption = "Velocidade do semáforo(segs): " & TEMPO_SEMAFORO / 1000
    TEMPO_SINAL = TEMPO_SEMAFORO / 2
    
    frmSimC.tsemA.Interval = TEMPO_SINAL
    frmSimC.tsemB.Interval = TEMPO_SEMAFORO
    
    frmSimD.tsemA.Interval = TEMPO_SINAL
    frmSimD.tsemB.Interval = TEMPO_SEMAFORO
    
End Sub
