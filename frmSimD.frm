VERSION 5.00
Begin VB.Form frmSimD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Simulação 2 pistas semáforo 3 tempos"
   ClientHeight    =   6765
   ClientLeft      =   6915
   ClientTop       =   315
   ClientWidth     =   6765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSimD.frx":0000
   ScaleHeight     =   6765
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Tick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   5880
      Top             =   915
   End
   Begin VB.Timer geraCarros 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   5895
      Top             =   1410
   End
   Begin VB.Timer tsemA 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5400
      Top             =   915
   End
   Begin VB.Timer tsemB 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   5385
      Top             =   1410
   End
   Begin VB.Image verde_direita4 
      Height          =   165
      Left            =   3990
      Picture         =   "frmSimD.frx":948D2
      Top             =   2235
      Width           =   195
   End
   Begin VB.Image verde_direita3 
      Height          =   165
      Left            =   2550
      Picture         =   "frmSimD.frx":94ACC
      Top             =   3885
      Width           =   195
   End
   Begin VB.Image verde_direita2 
      Height          =   195
      Left            =   2460
      Picture         =   "frmSimD.frx":94CC6
      Top             =   2325
      Width           =   165
   End
   Begin VB.Image verde_direita1 
      Height          =   195
      Left            =   4110
      Picture         =   "frmSimD.frx":94EDC
      Top             =   3765
      Width           =   165
   End
   Begin VB.Image imgCarro 
      Height          =   225
      Index           =   0
      Left            =   375
      Picture         =   "frmSimD.frx":950F2
      Top             =   270
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgSemD 
      Height          =   465
      Left            =   3810
      Picture         =   "frmSimD.frx":95404
      Top             =   2235
      Width           =   195
   End
   Begin VB.Image imgSemC 
      Height          =   465
      Left            =   2730
      Picture         =   "frmSimD.frx":9591E
      Top             =   3585
      Width           =   195
   End
   Begin VB.Image imgSemB 
      Height          =   195
      Left            =   2460
      Picture         =   "frmSimD.frx":95E38
      Top             =   2505
      Width           =   465
   End
   Begin VB.Image imgSemA 
      Height          =   195
      Left            =   3810
      Picture         =   "frmSimD.frx":9635A
      Top             =   3585
      Width           =   465
   End
End
Attribute VB_Name = "frmSimD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mapa(1 To 30, 1 To 30) As tNodo

Private numCarro As Long

Private TOTAL_CARROS_E As Long
Private TOTAL_CARROS_S As Long

Private carros() As tCarro
Private tmpCarros() As tCarro
 
Private semA As Boolean
Private sinA As Long

Private semB As Boolean
Private sinB As Long

Private semC As Boolean
Private sinC As Long

Private semD As Boolean
Private sinD As Long

Private arquivo As clsArquivo

Public Sub start()
    
    Set arquivo = New clsArquivo
    arquivo.Nome = "Sim_3Tempos.csv.txt"
    arquivo.Diretorio = App.Path
    arquivo.abreArquivo (F_WRITE)
    arquivo.writeLine ("Carro ID;Origem (x,y);Destino (x,y);Entrada;Saida;Total")
    
    semA = False
    sinA = semaforo.VERMELHO
    semB = True
    sinB = semaforo.VERDE
    
    tsemA.Enabled = True
    tsemB.Enabled = True
    
    inicializaSemaforos
    atualizaSemaforos
    
    Tick.Enabled = True
    geraCarros.Enabled = True
    tsemA.Enabled = True
    tsemB.Enabled = True
    
End Sub

Public Sub pause()
    
    Tick.Enabled = False
    geraCarros.Enabled = False
    tsemA.Enabled = False
    tsemB.Enabled = False

End Sub

Public Sub continue()
    
    Tick.Enabled = True
    geraCarros.Enabled = True
    tsemA.Enabled = True
    tsemB.Enabled = True
    
End Sub

Public Sub stops()

    Dim i As Long
    
    Tick.Enabled = False
    geraCarros.Enabled = False
    tsemA.Enabled = False
    tsemB.Enabled = False
    arquivo.fClose
    
    For i = 1 To imgCarro.Count - 1
        Unload imgCarro(i)
    Next i
    
    numCarro = 0
    desocupaMapa
    TOTAL_CARROS_E = 0
    TOTAL_CARROS_S = 0
End Sub

Private Sub geraCarros_Timer()
    
    Dim i As Long, j As Long
    Dim tmpCarros() As tCarro
    
    Dim carro As tCarro
    geraCarros.Enabled = False
    If (TOTAL_CARROS_E - TOTAL_CARROS_S) < MAX_CARROS Then
        numCarro = numCarro + 1
        
        Load imgCarro(numCarro)
        incializaCarro carro, imgCarro(numCarro), numCarro
        posicionaCarroMapa carro, proximaPosicao(carro, carro.incremento)
        carro.entrada = Time
        If numCarro > 1 Then
            recriaVetorCarros
            ReDim Preserve carros(1 To UBound(carros) + 1)
            carros(UBound(carros)) = carro
        Else
            ReDim carros(1)
            carros(1) = carro
        End If
        
        TOTAL_CARROS_E = TOTAL_CARROS_E + 1
    End If
    geraCarros.Enabled = True
End Sub


Private Sub Tick_Timer()
    
    Dim saida As Date
    Dim carro As tCarro
    Dim pos As tPonto
    Dim inc As tPonto
    Dim i As Long
    
    Tick.Enabled = False
    
    If numCarro > 0 Then
        
        For i = 1 To UBound(carros) 'numCarro
            carro = carros(i)
            If Not chegouFim(carro) Then
                
                pos = proximaPosicao(carro, pegaIncremento(carro))
    
                'If estaEmCruzamento(carro) Then
                '    If adjacenciasOcupadas(carro) Then
                '        pos = proximaPosicao(carro, carro.incremento)
                '        carro.caminhoAlternativo = True
                '    End If
                '     If Not mapa(pos.x, pos.y).ocupado Then posicionaCarroMapa carro, pos
                'Else
                If Not (mapa(pos.x, pos.y).ocupado) And _
                    (Not mapa(carro.posicao.x, carro.posicao.y).nSemaforo.possui Or _
                    (mapa(carro.posicao.x, carro.posicao.y).nSemaforo.possui And _
                    mapa(carro.posicao.x, carro.posicao.y).nSemaforo.aberto)) _
                    Then
                    carro.incremento = pegaIncremento(carro)
                    posicionaCarroMapa carro, pos
                End If
                
            Else
                If Not carro.chegouFim Then
                    If Not carros(i).chegouFim Then TOTAL_CARROS_S = TOTAL_CARROS_S + 1
                    mapa(carro.posicao.x, carro.posicao.y).ocupado = False
                    imgCarro(carro.id).Visible = False
                    saida = Time
                    saida = Time
                    arquivo.writeLine (carro.id & ";(" & carro.origem.x & "," & carro.origem.y & ");(" _
                                        & carro.destino.x & "," & carro.destino.y & ");" _
                                        & carro.entrada & ";" & saida & ";" _
                                        & DateDiff("s", carro.entrada, saida))
                    carro.chegouFim = True
                End If
            End If
            carros(i) = carro
        Next i
        'geraCarros.Enabled = False
        'geraCarros.Enabled = True
    End If
    Tick.Enabled = True
    
End Sub

Private Function adjacenciasOcupadas(carro As tCarro) As Boolean

    Dim inc As tPonto
    Dim pos As tPonto
    
    inc = pegaIncremento(carro)
    pos = proximaPosicao(carro, pegaIncremento(carro))
    
    If (mapa(pos.x, pos.y).ocupado) Or _
        (mapa(pos.x + carro.incremento.x, pos.y + carro.incremento.y).ocupado) Or _
        (mapa(pos.x + inc.x, pos.y + inc.y).ocupado) Or _
        (mapa(pos.x + inc.x + carro.incremento.x, pos.y + inc.y + carro.incremento.y).ocupado) Then
        adjacenciasOcupadas = True
    Else
        adjacenciasOcupadas = False
    End If

End Function

Private Sub posicionaCarroMapa(carro As tCarro, pos As tPonto)
    mapa(carro.posicao.x, carro.posicao.y).ocupado = False
    mapa(pos.x, pos.y).ocupado = True
    carro.posicao = pos
    posicionaCarro carro
End Sub

Private Sub tsemA_Timer()
    
    tsemA.Enabled = False
        
    Select Case sinA
        Case semaforo.VERMELHO:
            If semA Then
                sinA = semaforo.VERDE
                Set imgSemA.picture = LoadPicture(App.Path & "\verde_1.bmp")
            End If
        Case semaforo.AMARELO:
            If Not semA Then
                sinA = semaforo.VERMELHO
                Set imgSemA.picture = LoadPicture(App.Path & "\vermelho_1.bmp")
            End If
        Case semaforo.VERDE:
            If semA Then
                sinA = semaforo.AMARELO
                Set imgSemA.picture = LoadPicture(App.Path & "\amarelo_1.bmp")
            End If
    End Select
        
    Select Case sinB
        Case semaforo.VERMELHO:
            If semB Then
                sinB = semaforo.VERDE
                Set imgSemB.picture = LoadPicture(App.Path & "\verde_2.bmp")
            End If
        Case semaforo.AMARELO:
            If Not semB Then
                sinB = semaforo.VERMELHO
                Set imgSemB.picture = LoadPicture(App.Path & "\vermelho_2.bmp")
            End If
        Case semaforo.VERDE:
            If semB Then
                sinB = semaforo.AMARELO
                Set imgSemB.picture = LoadPicture(App.Path & "\amarelo_2.bmp")
            End If
    End Select
    
    Select Case sinC
        Case semaforo.VERMELHO:
            If semC Then
                sinC = semaforo.VERDE
                Set imgSemC.picture = LoadPicture(App.Path & "\verde_3.bmp")
            End If
        Case semaforo.AMARELO:
            If Not semC Then
                sinC = semaforo.VERMELHO
                Set imgSemC.picture = LoadPicture(App.Path & "\vermelho_3.bmp")
            End If
        Case semaforo.VERDE:
            If semC Then
                sinC = semaforo.AMARELO
                Set imgSemC.picture = LoadPicture(App.Path & "\amarelo_3.bmp")
            End If
    End Select
    
    Select Case sinD
        Case semaforo.VERMELHO:
            If semD Then
                sinD = semaforo.VERDE
                Set imgSemD.picture = LoadPicture(App.Path & "\verde_4.bmp")
            End If
        Case semaforo.AMARELO:
            If Not semD Then
                sinD = semaforo.VERMELHO
                Set imgSemD.picture = LoadPicture(App.Path & "\vermelho_4.bmp")
            End If
        Case semaforo.VERDE:
            If semD Then
                sinD = semaforo.AMARELO
                Set imgSemD.picture = LoadPicture(App.Path & "\amarelo_4.bmp")
            End If
    End Select
   
    tsemA.Enabled = True
    End Sub

Private Sub tsemB_Timer()
    If semA Then
        semA = False
        semB = False
        semC = True
        semD = False

        
    ElseIf semB Then
        semA = False
        semB = False
        semC = False
        semD = True

    ElseIf semC Then
        semA = False
        semB = True
        semC = False
        semD = False

    Else
        semA = True
        semB = False
        semC = False
        semD = False

    End If
    
    tsemA_Timer
    atualizaSemaforos
    
End Sub

Public Sub inicializaSemaforos()
    'SEMAFORO A
    mapa(16, 17).nSemaforo.possui = True
    mapa(17, 17).nSemaforo.possui = True
    mapa(14, 12).nSemaforo.possui = True
    mapa(15, 12).nSemaforo.possui = True
    'SEMAFORO B
    mapa(13, 15).nSemaforo.possui = True
    mapa(13, 16).nSemaforo.possui = True
    mapa(18, 13).nSemaforo.possui = True
    mapa(18, 14).nSemaforo.possui = True


End Sub

Private Sub atualizaSemaforos()
    
    'mapa(16, 17).nSemaforo.aberto = semA
    'mapa(17, 17).nSemaforo.aberto = semA
    'mapa(14, 12).nSemaforo.aberto = semA
    'mapa(15, 12).nSemaforo.aberto = semA

    'mapa(13, 15).nSemaforo.aberto = semB
    'mapa(13, 16).nSemaforo.aberto = semB
    'mapa(18, 13).nSemaforo.aberto = semB
    'mapa(18, 14).nSemaforo.aberto = semB
    
    If semA Then
        Set verde_direita1.picture = LoadPicture(App.Path & "\verde_direita1.a.bmp")
        Set verde_direita2.picture = LoadPicture(App.Path & "\verde_direita2.f.bmp")
        Set verde_direita3.picture = LoadPicture(App.Path & "\verde_direita3.a.bmp")
        Set verde_direita4.picture = LoadPicture(App.Path & "\verde_direita4.f.bmp")
        mapa(16, 17).nSemaforo.aberto = True
        mapa(17, 17).nSemaforo.aberto = True
        mapa(14, 12).nSemaforo.aberto = False
        mapa(15, 12).nSemaforo.aberto = False
        mapa(13, 15).nSemaforo.aberto = False
        mapa(13, 16).nSemaforo.aberto = True
        mapa(18, 13).nSemaforo.aberto = False
        mapa(18, 14).nSemaforo.aberto = False
    ElseIf semB Then
        Set verde_direita1.picture = LoadPicture(App.Path & "\verde_direita1.f.bmp")
        Set verde_direita2.picture = LoadPicture(App.Path & "\verde_direita2.a.bmp")
        Set verde_direita3.picture = LoadPicture(App.Path & "\verde_direita3.f.bmp")
        Set verde_direita4.picture = LoadPicture(App.Path & "\verde_direita4.a.bmp")
        mapa(16, 17).nSemaforo.aberto = False
        mapa(17, 17).nSemaforo.aberto = False
        mapa(14, 12).nSemaforo.aberto = True
        mapa(15, 12).nSemaforo.aberto = True
        mapa(13, 15).nSemaforo.aberto = False
        mapa(13, 16).nSemaforo.aberto = False
        mapa(18, 13).nSemaforo.aberto = True
        mapa(18, 14).nSemaforo.aberto = False
    ElseIf semC Then
        Set verde_direita1.picture = LoadPicture(App.Path & "\verde_direita1.f.bmp")
        Set verde_direita2.picture = LoadPicture(App.Path & "\verde_direita2.a.bmp")
        Set verde_direita3.picture = LoadPicture(App.Path & "\verde_direita3.a.bmp")
        Set verde_direita4.picture = LoadPicture(App.Path & "\verde_direita4.f.bmp")
        mapa(16, 17).nSemaforo.aberto = False
        mapa(17, 17).nSemaforo.aberto = False
        mapa(14, 12).nSemaforo.aberto = True
        mapa(15, 12).nSemaforo.aberto = False
        mapa(13, 15).nSemaforo.aberto = True
        mapa(13, 16).nSemaforo.aberto = True
        mapa(18, 13).nSemaforo.aberto = False
        mapa(18, 14).nSemaforo.aberto = False
    Else
        Set verde_direita1.picture = LoadPicture(App.Path & "\verde_direita1.a.bmp")
        Set verde_direita2.picture = LoadPicture(App.Path & "\verde_direita2.f.bmp")
        Set verde_direita3.picture = LoadPicture(App.Path & "\verde_direita3.f.bmp")
        Set verde_direita4.picture = LoadPicture(App.Path & "\verde_direita4.a.bmp")
        mapa(16, 17).nSemaforo.aberto = False
        mapa(17, 17).nSemaforo.aberto = True
        mapa(14, 12).nSemaforo.aberto = False
        mapa(15, 12).nSemaforo.aberto = False
        mapa(13, 15).nSemaforo.aberto = False
        mapa(13, 16).nSemaforo.aberto = False
        mapa(18, 13).nSemaforo.aberto = True
        mapa(18, 14).nSemaforo.aberto = True
    End If


    
End Sub

Public Sub recriaVetorCarros()
    
    Dim i As Long, j As Long
   
    For i = 1 To UBound(carros)
        If Not carros(i).chegouFim Then
            j = j + 1
            ReDim Preserve tmpCarros(1 To j)
            tmpCarros(j) = carros(i)
        End If
    Next i
    carros = tmpCarros
    
End Sub

Public Sub desocupaMapa()

    Dim i As Integer, j As Integer

    For i = 1 To 30
        For j = 1 To 30
            mapa(i, j).ocupado = False
        Next j
    Next i
End Sub
