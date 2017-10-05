VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Nodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private pEsquerda As PictureBox
Private pDireita As PictureBox
Private pCima As PictureBox
Private pBaixo As PictureBox
Private pSelf As PictureBox

Public Property Set esquerda(aEsquerda As PictureBox)
    pEsquerda = aEsquerda
End Property

Public Property Get esquerda() As PictureBox
    esquerda = pEsquerda
End Property

Public Property Set direita(aDireita As PictureBox)
    pDireita = aDireita
End Property

Public Property Get direita() As PictureBox
    direita = pDireita
End Property

Public Property Set cima(aCima As PictureBox)
    pCima = aCima
End Property

Public Property Get cima() As PictureBox
    cima = pCima
End Property

Public Property Set baixo(aBaixo As PictureBox)
    pBaixo = aBaixo
End Property

Public Property Get baixo() As PictureBox
    baixo = pBaixo
End Property

Public Property Set self(aSelf As PictureBox)
    pSelf = aSelf
End Property

Public Property Get self() As PictureBox
    self = pSelf
End Property


