VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRoteador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public numero As Long

Private pOcupado As Boolean

Private pRotas As Collection

Private pPicture As PictureBox

Private pSemaforo   As Boolean
Private pAberto     As Boolean

Public Property Set rotas(aRota As Collection)
    Set pRotas = aRota
End Property

Public Property Get rotas() As Collection
    Set rotas = pRotas
End Property

Public Property Let ocupado(aOcupado As Boolean)
    pOcupado = aOcupado
End Property

Public Property Get ocupado() As Boolean
    ocupado = pOcupado
End Property

Public Property Set picture(aPicture As PictureBox)
    Set pPicture = aPicture
End Property

Public Property Get picture() As PictureBox
    Set picture = pPicture
End Property

Public Function getProximo(destino As Long, nodos As Variant) As Object
    
    Dim rota As clsRota
    
    For Each rota In pRotas
        If rota.destino = destino Then
            Set getProximo = rota.proximo
            If Not rota.proximo.ocupado Then Exit For
        End If
    Next rota
            
End Function

Private Sub Class_Initialize()
    Set pRotas = New Collection
End Sub

Public Property Let semaforo(aSemaforo As Boolean)
    pSemaforo = aSemaforo
End Property
Public Property Get semaforo() As Boolean
    semaforo = pSemaforo
End Property

Public Property Let aberto(aAberto As Boolean)
    pAberto = aAberto
End Property
Public Property Get aberto() As Boolean
    aberto = pAberto
End Property
