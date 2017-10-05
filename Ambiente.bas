Attribute VB_Name = "Ambiente"

Option Explicit

Public Type tPonto
    x As Integer
    y As Integer
End Type

Public Type tCarro
    origem              As tPonto
    destino             As tPonto
    posicao             As tPonto
    incremento          As tPonto
    caminhoAlternativo  As Boolean
    chegouFim           As Boolean
    id                  As Long
    img                 As Image
    entrada             As String
End Type

Public Type tSemaforo
    aberto As Boolean
    possui As Boolean
End Type

Public Type tNodo
    ocupado As Boolean
    nSemaforo As tSemaforo
End Type

Public Sub posicionaCarro(carro As tCarro)
    
    With carro
        .img.Top = (.posicao.y - 1) * .img.Height
        .img.Left = (.posicao.x - 1) * .img.Width
    End With
    
End Sub

Public Function pegaIncremento(carro As tCarro) As tPonto
    
    Dim incremento As tPonto
    incremento = carro.incremento
    With carro
        
        If Not .caminhoAlternativo Then
        
            If .posicao.y = .destino.y And incremento.x = 0 Then
                incremento.y = 0
                If .posicao.x < .destino.x Then
                    incremento.x = 1
                ElseIf .posicao.x > .destino.x Then
                    incremento.x = -1
                End If
            ElseIf .posicao.x = .destino.x And incremento.y = 0 Then
                incremento.x = 0
                If .posicao.y < .destino.y Then
                    incremento.y = 1
                ElseIf .posicao.y > .destino.y Then
                    incremento.y = -1
                End If
            End If
        
        Else
            If .origem.x = 16 And .origem.y = 26 Then
                If .posicao.x = 16 And .posicao.y = 3 Then
                    incremento.x = 1
                    incremento.y = 0
                ElseIf .posicao.x = 28 And .posicao.y = 3 Then
                    incremento.x = 0
                    incremento.y = 1
                ElseIf .posicao.x = 28 And .posicao.y = 14 Then
                    incremento.x = -1
                    incremento.y = 0
                End If
            ElseIf .origem.x = 15 And .origem.y = 5 Then
                If .posicao.x = 15 And .posicao.y = 28 Then
                    incremento.x = -1
                    incremento.y = 0
                ElseIf .posicao.x = 3 And .posicao.y = 28 Then
                    incremento.x = 0
                    incremento.y = -1
                ElseIf .posicao.x = 3 And .posicao.y = 15 Then
                    incremento.x = 1
                    incremento.y = 0
                End If
            ElseIf .origem.x = 26 And .origem.y = 14 Then
                If .posicao.x = 3 And .posicao.y = 14 Then
                    incremento.x = 0
                    incremento.y = -1
                ElseIf .posicao.x = 3 And .posicao.y = 3 Then
                    incremento.x = 1
                    incremento.y = 0
                ElseIf .posicao.x = 15 And .posicao.y = 3 Then
                    incremento.x = 0
                    incremento.y = 1
                End If
            ElseIf .origem.x = 5 And .origem.y = 15 Then
                If .posicao.x = 28 And .posicao.y = 15 Then
                    incremento.x = 0
                    incremento.y = 1
                ElseIf .posicao.x = 28 And .posicao.y = 28 Then
                    incremento.x = -1
                    incremento.y = 0
                ElseIf .posicao.x = 16 And .posicao.y = 28 Then
                    incremento.x = 0
                    incremento.y = -1
                End If
            End If
        End If
        
    End With
    
    pegaIncremento = incremento
    
End Function

Public Function estaEmCruzamento(carro As tCarro)
    
    With carro
        
        If (.origem.x = 16 And .origem.y = 26 And _
            .destino.x = 5 And .destino.y = 14 And _
            .posicao.x = 16 And .posicao.y = 14) Or _
            (.origem.x = 26 And .origem.y = 14 And _
            .destino.x = 15 And .destino.y = 26 And _
            .posicao.x = 15 And .posicao.y = 14) Or _
            (.origem.x = 15 And .origem.y = 5 And _
            .destino.x = 26 And .destino.y = 15 And _
            .posicao.x = 15 And .posicao.y = 15) Or _
            (.origem.x = 5 And .origem.y = 15 And _
            .destino.x = 16 And .destino.y = 5 And _
            .posicao.x = 16 And .posicao.y = 15) _
            Then
            estaEmCruzamento = True
        Else
             estaEmCruzamento = False
        End If
        'If (((.destino.x < .origem.x) And (.destino.y < .origem.y)) Or _
        '   ((.destino.x > .origem.x) And (.destino.y > .origem.y)) Or _
        '   ((.destino.x > .origem.x) And (.destino.y < .origem.y)) Or _
        '   ((.destino.x < .origem.x) And (.destino.y > .origem.y))) And _
        '   ((.destino.y = .posicao.y And .posicao.x = .origem.x) Or _
        '   ((.destino.x = .posicao.x And .posicao.y = .origem.y))) Then
        'If ((.destino.y = .posicao.y And .posicao.x = .origem.x) Or _
        '   ((.destino.x = .posicao.x And .posicao.y = .origem.y))) Then
        '    estaEmCruzamento = True
        'Else
        '    estaEmCruzamento = False
        'End If
    End With
    
End Function

Public Function proximaPosicao(carro As tCarro, incremento As tPonto) As tPonto
   
    proximaPosicao.x = carro.posicao.x + incremento.x
    proximaPosicao.y = carro.posicao.y + incremento.y

End Function

Public Function chegouFim(carro As tCarro)
    With carro
        If .posicao.x = .destino.x And .posicao.y = .destino.y Then
            chegouFim = True
        Else
            chegouFim = False
        End If
    End With
End Function
  

Public Sub setOrigem(carro As tCarro)

    Dim intOrigem As Integer
    
    intOrigem = Rand(1, 8)
    
    Select Case intOrigem
        Case 1:
            carro.origem.x = 17
            carro.origem.y = 26
            carro.incremento.x = 0
            carro.incremento.y = -1
        Case 2:
            carro.origem.x = 16
            carro.origem.y = 26
            carro.incremento.x = 0
            carro.incremento.y = -1
        Case 3:
            carro.origem.x = 5
            carro.origem.y = 16
            carro.incremento.x = 1
            carro.incremento.y = 0
        Case 4:
            carro.origem.x = 5
            carro.origem.y = 15
            carro.incremento.x = 1
            carro.incremento.y = 0
        Case 5:
            carro.origem.x = 26
            carro.origem.y = 14
            carro.incremento.x = -1
            carro.incremento.y = 0
        Case 6:
            carro.origem.x = 26
            carro.origem.y = 13
            carro.incremento.x = -1
            carro.incremento.y = 0
        Case 7:
            carro.origem.x = 14
            carro.origem.y = 5
            carro.incremento.x = 0
            carro.incremento.y = 1
        Case 8:
            carro.origem.x = 15
            carro.origem.y = 5
            carro.incremento.x = 0
            carro.incremento.y = 1
    End Select

End Sub

Public Sub setDestino(carro As tCarro)
    Dim intDestino As Integer
    With carro
        intDestino = Rand(1, 2)
        If .origem.x = 17 And .origem.y = 26 Then
            'If intDestino = 1 Then
                .destino.x = 26
                .destino.y = 16
            'Else
            '    .destino.x = 17
            '    .destino.y = 5
            'End If
        ElseIf .origem.x = 16 And .origem.y = 26 Then
            If intDestino = 1 Then
                .destino.x = 5
                .destino.y = 14
            Else
                .destino.x = 16
                .destino.y = 5
            End If
        ElseIf .origem.x = 5 And .origem.y = 16 Then
            'If intDestino = 1 Then
                .destino.x = 14
                .destino.y = 26
            'Else
            '    .destino.x = 26
            '    .destino.y = 16
            'End If
        ElseIf .origem.x = 5 And .origem.y = 15 Then
            If intDestino = 1 Then
                .destino.x = 16
                .destino.y = 5
            Else
                .destino.x = 26
                .destino.y = 15
            End If
        ElseIf .origem.x = 26 And .origem.y = 14 Then
            If intDestino = 1 Then
                .destino.x = 15
                .destino.y = 26
            Else
                .destino.x = 5
                .destino.y = 14
            End If
        ElseIf .origem.x = 26 And .origem.y = 13 Then
            'If intDestino = 1 Then
                .destino.x = 17
                .destino.y = 5
            'Else
            '    .destino.x = 5
            '    .destino.y = 13
            'End If
        ElseIf .origem.x = 14 And .origem.y = 5 Then
            'If intDestino = 1 Then
                .destino.x = 5
                .destino.y = 13
            'Else
            '    .destino.x = 14
            '    .destino.y = 26
            'End If
        ElseIf .origem.x = 15 And .origem.y = 5 Then
            If intDestino = 1 Then
                .destino.x = 15
                .destino.y = 26
            Else
                .destino.x = 26
                .destino.y = 15
            End If
        End If
    End With
        
End Sub

Public Sub setImagem(carro As tCarro, img As Image)
    
    Dim intCor As Integer
    intCor = Rand(1, 6)
    
    Set carro.img = img
    Set carro.img.picture = LoadPicture(App.Path & "\carro_" & intCor & ".bmp")
    carro.img.Visible = True
End Sub

Public Sub incializaCarro(carro As tCarro, img As Image, id As Long)
    
    setOrigem carro
    setDestino carro
    setImagem carro, img
    
    carro.posicao = carro.origem
    
    carro.id = id
    
End Sub

