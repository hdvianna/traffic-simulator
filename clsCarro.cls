VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private pDestino As Long

Private pOrigem As Long

Private pNodo As Object

Private pcarId As String

Private pNum   As Long

Private pEntrada As Date
Private pSaida As Date

Private Sub geraOrigem()
    
    Dim idxOrigem As Integer
    
    idxOrigem = Rand(1, 4)
    
    Select Case idxOrigem
        Case 1: pOrigem = 30
        Case 2: pOrigem = 47
        Case 3: pOrigem = 117
        Case 4: pOrigem = 136
    End Select
    
End Sub

Private Sub geraDestino()
    
    Dim idxDestino As Integer
    
    idxDestino = Rand(1, 3)
    
    Select Case pOrigem
        Case 30:
            Select Case idxDestino
                Case 1: pDestino = 17
                Case 2: pDestino = 122
                Case 3: pDestino = 131
            End Select
        Case 47:
            Select Case idxDestino
                Case 1: pDestino = 60
                Case 2: pDestino = 122
                Case 3: pDestino = 131
            End Select
        Case 117:
            Select Case idxDestino
                Case 1: pDestino = 17
                Case 2: pDestino = 60
                Case 3: pDestino = 131
            End Select
        Case 136:
            Select Case idxDestino
                Case 1: pDestino = 17
                Case 2: pDestino = 60
                Case 3: pDestino = 122
            End Select
    End Select

End Sub

Public Property Set nodo(aNodo As Object)
    Set pNodo = aNodo
End Property

Public Property Get nodo() As Object
    Set nodo = pNodo
End Property

Public Property Let destino(aDestino As Long)
    pDestino = aDestino
End Property

Public Property Get destino() As Long
    destino = pDestino
End Property

Public Property Let origem(aOrigem As Long)
    pOrigem = aOrigem
End Property

Public Property Get origem() As Long
    origem = pOrigem
End Property

Public Property Let carid(acarId As String)
    pcarId = acarId
End Property

Public Property Get carid() As String
    carid = pcarId
End Property


Public Property Let entrada(aEntrada As Date)
    pEntrada = aEntrada
End Property

Public Property Get entrada() As Date
    entrada = pEntrada
End Property

Public Property Let saida(aSaida As Date)
    pSaida = aSaida
End Property

Public Property Get saida() As Date
    saida = pSaida
End Property

Public Property Get num() As Long
    num = pNum
End Property

Public Function chegou() As Boolean
    
    If pDestino = pNodo.picture.Index Then
        chegou = True
    Else
        chegou = False
    End If
        
End Function

Private Sub Class_Initialize()
    
    pNum = Rand(1, 6)
    geraOrigem
    geraDestino
    
End Sub
