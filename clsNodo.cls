VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private pProximo As Object

Private pOcupado As Boolean

Private pPicture As PictureBox

Private pSemaforo   As Boolean
Private pAberto     As Boolean

Public Property Set proximo(aProximo As Object)
    Set pProximo = aProximo
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
Public Function getProximo(destino As Long, nodo As Variant) As Object
    Set getProximo = pProximo
End Function

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


