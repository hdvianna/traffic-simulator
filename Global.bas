Attribute VB_Name = "Global"

Option Explicit

Public Enum semaforo
    VERMELHO = 0
    AMARELO = 1
    VERDE = 2
End Enum

Public TEMPO_SEMAFORO As Long
Public TEMPO_SINAL As Long

Public MAX_CARROS As Long  'Total de carros permitidos

Public Function Rand(ByVal Low As Long, _
                     ByVal High As Long) As Long
  Randomize
  Rand = Int((High - Low + 1) * Rnd) + Low
End Function

