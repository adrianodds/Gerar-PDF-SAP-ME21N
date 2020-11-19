'VARIÁVEIS GLOBAIS PARA TODAS AS FUNÇÕES
Dim pedido As String
Dim link As Variant
Dim usuario
Sub pdf()
  
'*********** CONEXÃO SAP ***************'

  Dim SapGuiAuto As Object
  Dim Application As SAPFEWSELib.GuiApplication
  Set SapGuiAuto = GetObject("SAPGUI")
  Set Application = SapGuiAuto.GetScriptingEngine
  If Not IsObject(Connection) Then
  On Error GoTo sap_error
   Set Connection = Application.Children(0)
   End If
   If Not IsObject(session) Then
   Set session = Connection.Children(0)
   End If
   If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
   End If
   
'*********** CONEXÃO SAP ***************'
Dim cont As Integer

cont = 2

'VERIFICANDO QUANTIDADES DE LINHAS A SEREM PERCORRIDAS

Do Until Range("A" & cont).Value = ""

cont = cont + 1

Loop

cont = cont - 1

For a = 2 To cont

pedido = Range("A" & a).Value

'VERIFICA SE JÁ ESTÁ O STATUS 'FEITO'
If Range("E" & a).Value = "" Then

'PREVIZUALIZAÇÃO EM PDF NO SAP
'obs: Essa previzualização irá gerar um arquivo temporário em pdf no seu computador

session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME23N"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").Text = pedido
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[20]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[0]/okcd").Text = "PDF!"
session.findById("wnd[0]").sendVKey 0

'CHAMANDO A SUB QUE VAI NA PASTA DE ARQUIVOS TEMPORÁRIOS
Call listar_arquivos_e_pastas

Range("E" & a).Value = "FEITO"

End If

Next

'QUANDO CONCLUIR O PROCESSO A PASTA SERÁ EXIBIDA
Shell "C:\WINDOWS\explorer.exe """ & "C:\Users\" & usuario & "\Desktop\ARQUIVOS PDF\" & "", vbNormalFocus

session.findById("wnd[0]/tbar[0]/okcd").Text = "/N"
session.findById("wnd[0]").sendVKey 0

'Fim do programa desejado---------------------------------------------------------------------------------------------------

    Set Connection = Nothing

Exit Sub

sap_error:
MsgBox "Certifique-se de estar logado ao SAP para prosseguir!", vbExclamation

End Sub

Sub listar_arquivos_e_pastas()

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim linha As Integer
    i = 1
    
    While Range("C" & i).Value <> ""
    
    i = i + 1
    
    Wend
    
    Set ObjNetwork = CreateObject("WScript.Network")
    usuario = ObjNetwork.UserName
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'CHAMANDO A SUB QUE IRÁ VERIFICAR SE SERÁ NECESSÁRIO MODIFICAR O USUÁRIO
    'obs: essa ação foi necessária pois na empresa em que trabalho alguns computadores tem usuários com uma informação adicional.
    Call user
    
    If link = "NORDESTAO" Then
    
    usuario = usuario & ".NORDESTAO"
    
    End If
    
    'ABRINDO PASTA COM OS ARQUIVOS TEMPORÁRIOS
    Set folder = fso.GetFolder("C:\Users\" & usuario & "\AppData\Local\Temp")
    
    'VERIFICANDO OS ARQUIVOS COM EXTENSÃO PDF E FILTRANDO PELOS QUE FORAM GERADOS HOJE
    For Each file In folder.Files
        
        extensao = Right(file.Name, 3)
        datas = Format(file.DateCreated, "DD.MM.YYYY")
        horas = Format(file.DateCreated, "H:M")
                
        If extensao = "pdf" Then
                
        If datas = Format(Now, "DD.MM.YYYY") Then
                
        Cells(i, 2) = file.Name
        Cells(i, 3) = datas
        Cells(i, 4) = horas
        
        On Error GoTo criar_pasta
        
        FileCopy "C:\Users\" & usuario & "\AppData\Local\Temp\" & file.Name, "C:\Users\" & usuario & "\Desktop\ARQUIVOS PDF\" & pedido & "_" & datas & "_" & Format(horas, "HH.MM") & ".pdf"
        
        End If
        
        End If
        
    Next file
    
    Exit Sub
    
criar_pasta:
MkDir "C:\Users\" & usuario & "\Desktop\ARQUIVOS PDF\"

Resume
End Sub

Sub limpar()

Range("A2", "F200").ClearContents

End Sub

Sub user()

link = ThisWorkbook.FullNameURLEncoded & "\"

link = Split(link, "\")

link = link(2)

link = Right(link, 9)

End Sub


