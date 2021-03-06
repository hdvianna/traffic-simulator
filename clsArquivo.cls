VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'************************************************************
'* CLASSE: clsArquivo                                       *
'* Descri��o: Esta classe abstrai o funcionamento de um     *
'*                arquivo texto.                            *
'************************************************************
Option Explicit
    
    Private pName       As String   'Onde ficar� o nome do arquivo
    Private pFolder     As String   'armazenar� o diret�rio do arquivo
    Private pNumber     As Integer  'armazenar� o n�mero do arquivo aberto
    Private pEndFile    As Boolean  'armazenar� true se o arquivo chegou ao fim
    Private pNumLinha   As Long     'armazenar� numero da linha corrente
    Private pTamArquivo As Long     'armazenar� o tamanho do arquivo
    
    Public Enum OPEN_FILE_MODE_ENUM 'Enumera��o das op��es para abertura de um arquivo
        F_READ = 0   'Arquivo de leitura, input
        F_WRITE = 1  'Arquivo de escrita, output
        F_APPEND = 2 'Arquivo de adjacente(?)
    End Enum

'######################
'# PROPRIEDADE Nome  ##
'######################
Public Property Let Nome(Value As String)
    'Atribui o nome do arquivo
    pName = Value
End Property
Public Property Get Nome() As String
    'Retorna o nome do arquivo
    Nome = pName
End Property

'###########################
'# PROPRIEDADE Diretorio  ##
'###########################
Public Property Let Diretorio(Value As String)
    'Atribui o diretorio onde ficar� o arquivo
    pFolder = Value
End Property
Public Property Get Diretorio() As String
    'Retorna o diretorio onde ficar� o arquivo
    Diretorio = pFolder
End Property

'#########################
'# PROPRIEDADE Caminho  ##
'#########################
Public Property Let Caminho(Value As String)
    'Atribui o caminho, contendo o nome do arquivo
    Dim intPos As Integer
    
    intPos = InStrRev(Value, "\") 'Pega a posi��o da barra que separa o diret�rio do nome do arquivo
    pFolder = Mid(Value, 1, intPos - 1) ' pega o diret�rio do caminho
    pName = Mid(Value, intPos + 1) 'pega o nome do arquivo
    
End Property
Public Property Get Caminho() As String
    Caminho = pFolder & "\" & pName
End Property

'##########################
'# PROPRIEDADE numLinha  ##
'##########################
Public Property Get numLinha() As Long
    'Retorna a posicao da linha lida
    numLinha = pNumLinha
End Property

'############################
'# PROPRIEDADE tamArquivo  ##
'############################
Public Property Get tamArquivo() As Long
    'Retorna o tamanho do arquivo
    tamArquivo = pTamArquivo
End Property

'#########################
'# M�TODO abreArquivo   ##
'#########################
'Abre um arquivo de texto de acordo com o modo passado
'como par�metro
Public Function abreArquivo(modo As OPEN_FILE_MODE_ENUM) As Boolean
    
On Error GoTo Erro
    
    pNumber = FreeFile 'pega o primeiro arquivo livre
    
    Select Case modo
        Case OPEN_FILE_MODE_ENUM.F_APPEND 'Abre o arquivo para adi��o
            Open pFolder & "\" & pName For Append As #pNumber
        Case OPEN_FILE_MODE_ENUM.F_READ 'Abre o arquivo para leitura
            Open pFolder & "\" & pName For Input As #pNumber
        Case OPEN_FILE_MODE_ENUM.F_WRITE 'Abre o arquivo pra escrita
            Open pFolder & "\" & pName For Output As #pNumber
    End Select
    
    abreArquivo = True 'Arquivo aberto com sucesso
    
    pNumLinha = 0 'Seta linha corrente para zero
    pTamArquivo = FileLen(pFolder & "\" & pName) ' Pega o tamanho do arquivo
    
    If EOF(pNumber) Then 'Verifica se o arquivo est� em seu final
        pEndFile = True
    Else
        pEndFile = False
    End If
    
    Exit Function

Erro:
    abreArquivo = False
    
End Function

'#####################
'# M�TODO getLine   ##
'#####################
'Este m�todo pega a linha corrente do arquivo texto
'Quando a linha a lida o cursor do arquivo vai pra a pr�xima linha
Public Function getLine() As String

    Dim strLine As String 'Armazenar� a linha do arquivo

On Error GoTo Erro
    
    Line Input #pNumber, strLine 'Pega a linha corrente do arquivo
    pNumLinha = pNumLinha + 1 'Incrementa o numero de linhas lidas
    
    getLine = strLine 'Retorna a linha lida
    
    If EOF(pNumber) Then 'Verifica se o arquivo chegou ao seu final
        pEndFile = True
    Else
        pEndFile = False
    End If
    
    Exit Function

Erro:
    getLine = ""

End Function

'#######################
'# M�TODO getEndFile  ##
'#######################
'Este m�todo retorna o valor da v�riavel que controla o final do arquivo
Public Function getEndfile() As Boolean

    getEndfile = pEndFile 'Retorna se o arquivo esta no fim

End Function


'#######################
'# M�TODO writeLine   ##
'#######################
'Escreve uma linha no arquivo, para isso o arquivo
'deve ser aberto como escrita
Public Function writeLine(strLinha As String) As Boolean

On Error GoTo Erro
    
    Print #pNumber, strLinha 'Escreve a linha no arquivo
    pNumLinha = pNumLinha + 1 'Acrescenta o numero de linhas lidas
    
    writeLine = True 'Linha escrita com sucesso
    Exit Function

Erro:
    writeLine = False '}=0 ... Algo estranho aconteceu

End Function

'####################
'# M�TODO fClose   ##
'####################
'Fecha o arquivo que estava aberto
Public Function fClose() As Boolean
    
On Error GoTo Erro

    Close #pNumber 'Fecha o arquivo corrente
    fClose = True  ' Arquivo fechado ...
    
    pNumLinha = 0   'Seta linha corrente para zero
    pTamArquivo = 0 'Seta o tamanho do arquivo para zero

    Exit Function
    
Erro:
    fClose = False


End Function

