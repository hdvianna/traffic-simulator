VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private pDestino As Long
Private pProximo As Object

Public Property Set proximo(aProximo As Object)
    Set pProximo = aProximo
End Property

Public Property Get proximo() As Object
    Set proximo = pProximo
End Property

Public Property Let destino(aDestino As Long)
    pDestino = aDestino
End Property

Public Property Get destino() As Long
    destino = pDestino
End Property


