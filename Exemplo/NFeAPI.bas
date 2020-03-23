Attribute VB_Name = "NFeAPI"
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'activate Microsoft XML, v6.0 in references

Private Const tempoResposta = 500
Private Const token = "4EB15D6DEDAEBAE3FD0B7B5E5B0AD6D4"

Function enviaConteudoParaAPI(conteudo As String, url As String, tpConteudo As String) As String
On Error GoTo SAI
    Dim contentType As String
    
    If (tpConteudo = "txt") Then
        contentType = "text/plain;charset=utf-8"
    ElseIf (tpConteudo = "xml") Then
        contentType = "application/xml;charset=utf-8"
    Else
        contentType = "application/json;charset=utf-8"
    End If
    
    Dim obj As MSXML2.ServerXMLHTTP60
    Set obj = New MSXML2.ServerXMLHTTP60
    obj.Open "POST", url
    obj.setRequestHeader "Content-Type", contentType
    If Trim(token) <> "" Then
        obj.setRequestHeader "X-AUTH-TOKEN", token
    End If
    obj.send conteudo
    Dim resposta As String
    resposta = obj.responseText
    
    Select Case obj.status
        Case 401
            MsgBox ("Token não enviado ou inválido")
        Case 403
            MsgBox ("Token sem permissão")
    End Select
    
    enviaConteudoParaAPI = resposta
    Exit Function
SAI:
  enviaConteudoParaAPI = "{" & """status"":""" & Err.Number & """," & """motivo"":""" & Err.Description & """" & "}"
End Function

'Esta função realiza o processo completo de emissão: envio, consulta e download do documento
Public Function emitirNFeSincrono(conteudo As String, tpConteudo As String, cnpj As String, tpDown As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
    Dim retorno As String
    Dim resposta As String
    Dim statusEnvio As String
    Dim statusConsulta As String
    Dim statusDownload As String
    Dim motivo As String
    Dim erros As String
    Dim nsNRec As String
    Dim chNFe As String
    Dim cStat As String
    Dim nProt As String

    statusEnvio = ""
    statusConsulta = ""
    statusDownload = ""
    motivo = ""
    erros = ""
    nsNRec = ""
    chNFe = ""
    cStat = ""
    nProt = ""
    
    gravaLinhaLog ("[EMISSAO_SINCRONA_INICIO]")
    
    resposta = emitirNFe(conteudo, tpConteudo)
    statusEnvio = LerDadosJSON(resposta, "status", "", "")

    If (statusEnvio = "200") Or (statusEnvio = "-6") Then
    
        nsNRec = LerDadosJSON(resposta, "nsNRec", "", "")

        Sleep (tempoResposta)

        resposta = consultarStatusProcessamento(cnpj, nsNRec, tpAmb)
        statusConsulta = LerDadosJSON(resposta, "status", "", "")

        If (statusConsulta = "200") Then
            
            cStat = LerDadosJSON(resposta, "cStat", "", "")

            If (cStat = "100") Or (cStat = "150") Then
            
                chNFe = LerDadosJSON(resposta, "chNFe", "", "")
                nProt = LerDadosJSON(resposta, "nProt", "", "")
                motivo = LerDadosJSON(resposta, "xMotivo", "", "")

                resposta = downloadNFeESalvar(chNFe, tpAmb, tpDown, caminho, exibeNaTela)
                statusDownload = LerDadosJSON(resposta, "status", "", "")
                
                If (statusDownload <> "200") Then
                
                    motivo = LerDadosJSON(resposta, "motivo", "", "")
                    
                End If
            Else
            
                motivo = LerDadosJSON(resposta, "xMotivo", "", "")
                
            End If
        ElseIf (statusConsulta = "-2") Then
                
            erros = Split(resposta, """erro"":""")
            erros = LerDadosJSON(resposta, "erro", "", "")
            motivo = LerDadosJSON(erros, "xMotivo", "", "")
            cStat = LerDadosJSON(erros, "cStat", "", "")
                        
        Else
        
            motivo = LerDadosJSON(resposta, "motivo", "", "")
            
        End If
        
    ElseIf (status = "-7") Then
    
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        nsNRec = LerDadosJSON(resposta, "nsNRec", "", "")
   
    ElseIf (statusEnvio = "-4") Or (statusEnvio = "-2") Then

        motivo = LerDadosJSON(resposta, "motivo", "", "")
        erros = LerDadosJSON(resposta, "erros", "", "")

    ElseIf (statusEnvio = "-999") Or (statusEnvio = "-5") Then
    
        erros = Split(resposta, """erro"":""")
        erros = LerDadosJSON(resposta, "erro", "", "")
        erros = LerDadosJSON(erros, "xMotivo", "", "")
        
    Else
    
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        
    End If
    
    'Monta o JSON de retorno
    retorno = "{"
    retorno = retorno & """statusEnvio"":""" & statusEnvio & ""","
    retorno = retorno & """statusConsulta"":""" & statusConsulta & ""","
    retorno = retorno & """statusDownload"":""" & statusDownload & ""","
    retorno = retorno & """cStat"":""" & cStat & ""","
    retorno = retorno & """chNFe"":""" & chNFe & ""","
    retorno = retorno & """nProt"":""" & nProt & ""","
    retorno = retorno & """motivo"":""" & motivo & ""","
    retorno = retorno & """nsNRec"":""" & nsNRec & ""","
    retorno = retorno & """erros"":""" & erros & """"
    retorno = retorno & "}"
    
    gravaLinhaLog ("[JSON_RETORNO]")
    gravaLinhaLog (retorno)
    gravaLinhaLog ("[EMISSAO_SINCRONA_FIM]")

    emitirNFeSincrono = retorno
End Function

Public Function emitirNFe(conteudo As String, tpConteudo As String) As String

    Dim url As String
    Dim resposta As String

    url = "https://nfe.ns.eti.br/nfe/issue"

    gravaLinhaLog ("[ENVIO_DADOS]")
    gravaLinhaLog (conteudo)
        
    resposta = enviaConteudoParaAPI(conteudo, url, tpConteudo)
    
    gravaLinhaLog ("[ENVIO_RESPOSTA]")
    gravaLinhaLog (resposta)

    emitirNFe = resposta
End Function

Public Function consultarStatusProcessamento(cnpj As String, nsNRec As String, tpAmb As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """CNPJ"":""" & cnpj & ""","
    json = json & """nsNRec"":""" & nsNRec & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"

    url = "https://nfe.ns.eti.br/nfe/issue/status"
    
    gravaLinhaLog ("[CONSULTA_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CONSULTA_RESPOSTA]")
    gravaLinhaLog (resposta)

    consultarStatusProcessamento = resposta
End Function

Public Function downloadNFe(chNFe As String, tpDown As String, tpAmb As String) As String

    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String

    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """tpDown"":""" & tpDown & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"

    url = "https://nfe.ns.eti.br/nfe/get"

    gravaLinhaLog ("[DOWNLOAD_NFE_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    status = LerDadosJSON(resposta, "status", "", "")
        
    If (status <> "200") Then
    
        gravaLinhaLog ("[DOWNLOAD_NFE_RESPOSTA]")
        gravaLinhaLog (resposta)
        
    Else

        gravaLinhaLog ("[DOWNLOAD_NFE_STATUS]")
        gravaLinhaLog (status)
        
    End If

    downloadNFe = resposta
End Function

Public Function downloadNFeESalvar(chNFe As String, tpAmb As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String

    Dim baixarXML As Boolean
    Dim baixarPDF As Boolean
    Dim baixarJSON As Boolean
    Dim xml As String
    Dim json As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String

    resposta = downloadNFe(chNFe, tpDown, tpAmb)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "200" Then
    
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
        
        If InStr(1, tpDown, "X") Then
        
            xml = LerDadosJSON(resposta, "xml", "", "")
            Call salvarXML(xml, caminho, chNFe, "", "")
            
        End If
        
        
        If InStr(1, tpDown, "P") Then
        
            pdf = LerDadosJSON(resposta, "pdf", "", "")
            Call salvarPDF(pdf, caminho, chNFe, "", "")
            
            If exibeNaTela Then
            
                ShellExecute 0, "open", caminho & chNFe & "-procNFe.pdf", "", "", vbNormalFocus
            
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informacoes")
    End If

    downloadNFeESalvar = resposta
End Function

Public Function downloadEventoNFe(chNFe As String, tpAmb As String, tpDown As String, tpEvento As String, nSeqEvento As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    If (tpEvento <> "INUT") Then
    
        'Monta o JSON
        json = "{"
        json = json & """chNFe"":""" & chNFe & ""","
        json = json & """tpAmb"":""" & tpAmb & ""","
        json = json & """tpDown"":""" & tpDown & ""","
        json = json & """tpEvento"":""" & tpEvento & ""","
        json = json & """nSeqEvento"":""" & nSeqEvento & """"
        json = json & "}"
        
        url = "https://nfe.ns.eti.br/nfe/get/event"
    Else
        json = "{"
        json = json & """chave"":""" & chNFe & ""","
        json = json & """tpAmb"":""" & tpAmb & ""","
        json = json & """tpDown"":""" & tpDown & """"
        json = json & "}"
        
        url = "https://nfe.ns.eti.br/nfe/get/inut"
    End If
        
    
    gravaLinhaLog ("[DOWNLOAD_EVENTO_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")

    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status <> "200") Then
        gravaLinhaLog ("[DOWNLOAD_EVENTO_RESPOSTA]")
        gravaLinhaLog (resposta)
    Else
        gravaLinhaLog ("[DOWNLOAD_EVENTO_STATUS]")
        gravaLinhaLog (status)
    End If

    downloadEventoNFe = resposta
End Function

Public Function downloadEventoNFeESalvar(chNFe As String, tpAmb As String, tpDown As String, tpEvento As String, nSeqEvento As String, caminho As String, exibeNaTela As Boolean) As String
    Dim baixarXML As Boolean
    Dim baixarPDF As Boolean
    Dim baixarJSON As Boolean
    Dim xml As String
    Dim json As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    
    resposta = downloadEventoNFe(chNFe, tpAmb, tpDown, tpEvento, nSeqEvento)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "200" Then
        
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
        
        If InStr(1, tpDown, "X") Then
            If (tpEvento = "INUT") Then
                xml = LerDadosJSON(resposta, "retInut", "xml", "")
            Else
                xml = LerDadosJSON(resposta, "xml", "", "")
            End If
            
            Call salvarXML(xml, caminho, chNFe, tpEvento, nSeqEvento)
            
        End If

        If InStr(1, tpDown, "P") Then
            If (tpEvento = "INUT") Then
                pdf = LerDadosJSON(resposta, "retInut", "pdf", "")
            Else
                pdf = LerDadosJSON(resposta, "pdf", "", "")
            End If
            Call salvarPDF(pdf, caminho, chNFe, tpEvento, nSeqEvento)
            
            If exibeNaTela Then
    
                ShellExecute 0, "open", caminho & tpEvento & chNFe & nSeqEvento & "-procEvenNFe.pdf", "", "", vbNormalFocus
            
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informações")
    End If

    downloadEventoNFeESalvar = resposta
End Function

Public Function cancelarNFe(chNFe As String, tpAmb As String, dhEvento As String, nProt As String, xJust As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String
    
    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """dhEvento"":""" & dhEvento & ""","
    json = json & """nProt"":""" & nProt & ""","
    json = json & """xJust"":""" & xJust & """"
    json = json & "}"
    
    url = "https://nfe.ns.eti.br/nfe/cancel"
    
    gravaLinhaLog ("[CANCELAMENTO_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(json, url, "json")
        
    gravaLinhaLog ("[CANCELAMENTO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    'Se houve sucesso no evento, realiza o download
    If (status = "200") Then
        respostaDownload = downloadEventoNFeESalvar(chNFe, tpAmb, tpDown, "CANC", "1", caminho, exibeNaTela)
        
        status = LerDadosJSON(respostaDownload, "status", "", "")
        
        If (status <> "200") Then
            MsgBox ("Ocorreu um erro ao fazer o download. Verifique os logs.")
        End If
    End If
    
    cancelarNFe = resposta
End Function

Public Function corrigirNFe(chNFe As String, tpAmb As String, dhEvento As String, nSeqEvento As String, xCorrecao As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String
    
    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """dhEvento"":""" & dhEvento & ""","
    json = json & """nSeqEvento"":""" & nSeqEvento & ""","
    json = json & """xCorrecao"":""" & xCorrecao & """"
    json = json & "}"
    
    url = "https://nfe.ns.eti.br/nfe/cce"
    
    gravaLinhaLog ("[CCE_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CCE_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    'Se houve sucesso no evento, realiza o download
    If (status = "200") Then
        respostaDownload = downloadEventoNFeESalvar(chNFe, tpAmb, tpDown, "CCE", nSeqEvento, caminho, exibeNaTela)
        
        status = LerDadosJSON(respostaDownload, "status", "", "")
        
        If (status <> "200") Then
            MsgBox ("Ocorreu um erro ao fazer o download. Verifique os logs.")
        End If
    End If
    
    corrigirNFe = resposta
End Function

Public Function consultarCadastroContribuinte(CNPJCont As String, UF As String, documentoConsulta As String, tpConsulta As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """CNPJCont"":""" & CNPJCont & ""","
    json = json & """UF"":""" & UF & ""","
    json = json & """" & tpConsulta & """:""" & documentoConsulta & """"
    json = json & "}"

    url = "https://nfe.ns.eti.br/util/conscad"
    
    gravaLinhaLog ("[CONSULTA_CADASTRO_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CONSULTA_CADASTRO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    consultarCadastroContribuinte = resposta
End Function

Public Function consultarSituacao(licencaCnpj As String, chNFe As String, tpAmb As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """licencaCnpj"":""" & licencaCnpj & ""","
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """tpAmb"":""" & tpAmb & """}"

    url = "https://nfe.ns.eti.br/nfe/stats"
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_DADOS]")
    gravaLinhaLog (json)

    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    consultarSituacao = resposta
End Function

Public Function enviarEmail(chNFe As String, enviaEmailDoc As String, email As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """enviaEmailDoc"":" & enviaEmailDoc & ","
    json = json & """email"":["
    
    Dim emails() As String
    Dim i, quantidade As Integer
    
    emails = Split(Trim(email), ",")
    
    quantidade = UBound(emails)
    
    For i = 0 To quantidade
        If (i = quantidade) Then
            json = json & """" & emails(i) & """"
        Else
            json = json & """" & emails(i) & ""","
        End If
    Next
    
    json = json & "]"
    json = json & "}"

    url = "https://nfe.ns.eti.br/util/resendemail"
    
    gravaLinhaLog ("[ENVIO_EMAIL_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")

    gravaLinhaLog ("[ENVIO_EMAIL_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    enviarEmail = resposta
End Function

Public Function inutilizar(cUF As String, tpAmb As String, tpDown As String, ano As String, cnpj As String, serie As String, nNFIni As String, nNFFin As String, xJust As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim respostaDownload As String
    Dim chave As String

    'Monta o JSON
    json = "{"
    json = json & """cUF"":""" & cUF & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """ano"":""" & ano & ""","
    json = json & """CNPJ"":""" & cnpj & ""","
    json = json & """serie"":""" & serie & ""","
    json = json & """nNFIni"":""" & nNFIni & ""","
    json = json & """nNFFin"":""" & nNFFin & ""","
    json = json & """xJust"":""" & xJust & """"
    json = json & "}"

    url = "https://nfe.ns.eti.br/nfe/inut"
    
    gravaLinhaLog ("[INUTILIZACAO_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[INUTILIZACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    cStat = LerDadosJSON(resposta, "retornoInutNFe", "cStat", "")
    
    If (status = "200") And (cStat = "102") Then
        chave = LerDadosJSON(resposta, "retornoInutNFe", "chave", "")
        respostaDownload = downloadEventoNFeESalvar(chave, tpAmb, tpDown, "INUT", "1", caminho, exibeNaTela)
    Else
        MsgBox ("Inutilização com problema no Download")
    End If
    inutilizar = resposta
End Function

Public Function listarNSNRecs(chNFe As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & """"
    json = json & "}"

    url = "https://nfe.ns.eti.br/util/list/nsnrecs"
    
    gravaLinhaLog ("[LISTA_NSNRECS_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[LISTA_NSNRECS_RESPOSTA]")
    gravaLinhaLog (resposta)

    listarNSNRecs = resposta
End Function

Public Function previaNFe(conteudo As String, tpConteudo As String) As String

    Dim url As String
    Dim resposta As String

    url = "https://nfe.ns.eti.br/util/preview/nfe"

    gravaLinhaLog ("[ENVIO_PREVIA_DADOS]")
    gravaLinhaLog (conteudo)
        
    resposta = enviaConteudoParaAPI(conteudo, url, tpConteudo)
    
    gravaLinhaLog ("[ENVIO_PREVIA_RESPOSTA]")
    gravaLinhaLog (resposta)

    previaNFe = resposta
End Function

Public Function previaNFeESalvar(conteudo As String, tpConteudo As String, caminho As String, nomeArquivo As String, exibeNaTela As Boolean) As String

    Dim resposta As String
    Dim status As String
    Dim pdf As String

    resposta = previaNFe(conteudo, tpConteudo)
    
    status = LerDadosJSON(resposta, "status", "", "")
    pdf = LerDadosJSON(resposta, "pdf", "", "")
    
    If (status = "200") Then
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
        Call salvarPDF(pdf, caminho, nomeArquivo, "previa", "")
        If exibeNaTela Then

            ShellExecute 0, "open", caminho & nomeArquivo & nSeqEvento & "-procNFe.pdf", "", "", vbNormalFocus
            
        End If
    Else
        MsgBox ("Ocorreu um erro ao fazer a requisicao de previa da NFe. Verifique os logs para obter mais informacoes.")
    End If

    previaNFeESalvar = resposta
End Function

Public Function cadastrarLicenca(cnpj As String, razao As String, fantasia As String, ie As String, tipoICMS As String, email As String, endereco As String, numero As String, bairro As String, cep As String, cIBGE As String, telefone As String, situacao As String, idprojeto As String, usarAssinaturaLocal As String, certificado As String, senhaCert As String, server As String, porta As String, ssl As String, confirmaLeitura As String, usuario As String, senhaUser As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim emails() As String
    Dim telefones() As String
    Dim i, quantidade As Integer
    
    json = "{"
    json = json & """licencas"":[{"
    json = json & """situacao"":""" & situacao & ""","
    json = json & """idprojeto"":""" & idprojeto & ""","
    json = json & """usarcertns"":" & usarAssinaturaLocal
    
    If (usarAssinaturaLocal = "false") Then
        json = json & "," & """certificado"":{"
        json = json & """certificado"":""" & certificado & ""","
        json = json & """senha"":""" & senhaCert & """}"
    End If
    
    If (server <> "") Then
        json = json & "," & """emailenvio"":{"
        json = json & """servidor"":""" & server & ""","
        json = json & """porta"":""" & porta & ""","
        json = json & """ssl"":""" & ssl & ""","
        json = json & """confirmaleitura"":" & confirmaLeitura & ","
        json = json & """usuario"":""" & usuario & ""","
        json = json & """senha"":""" & senhaUser & """}"
    End If
    
    json = json & "}],"
    
    'Nodo da pessoa fisica ou juridica
    json = json & """pessoa"":{"
    json = json & """cnpj"":""" & cnpj & ""","
    json = json & """razao"":""" & razao & ""","
    json = json & """ie"":""" & ie & ""","
    json = json & """fantasia"":""" & fantasia & ""","
    json = json & """tipoicms"":""" & tipoICMS & ""","
    json = json & """emails"":[{"
    emails = Split(Trim(email), ",")
    quantidade = UBound(emails)
    
    For i = 0 To quantidade
        If (i = quantidade) Then
            json = json & """email"":""" & emails(i) & """}],"
        Else
            json = json & """email"":""" & emails(i) & ""","
        End If
    Next
    
    json = json & """enderecos"":[{"
    json = json & """endereco"":""" & endereco & ""","
    json = json & """numero"":""" & numero & ""","
    json = json & """bairro"":""" & bairro & ""","
    json = json & """cep"":""" & cep & ""","
    json = json & """cidade"":{"
    json = json & """cIBGE"":" & cIBGE & "}}],"
    
    json = json & """telefones"":[{"
    telefones = Split(Trim(telefone), ",")
    quantidade = UBound(telefones)
    
    For i = 0 To quantidade
        If (i = quantidade) Then
            json = json & """numero"":""" & telefones(i) & """}]"
        Else
            json = json & """numero"":""" & telefones(i) & ""","
        End If
    Next
    
    json = json & "}}"
     
    url = "http://painelapi.ns.eti.br/licenca/salvarDados"
    
    gravaLinhaLog ("[CADASTRAR_LICENCA_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CADASTRAR_LICENCA_RESPOSTA]")
    gravaLinhaLog (resposta)

    cadastrarLicenca = resposta
End Function

Public Sub salvarXML(xml As String, caminho As String, chNFe As String, tpEvento As String, nSeqEvento As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim tpEventoSalvar As String
    Dim localParaSalvar As String
    
    If (nSeqEvento = "") Then
        localParaSalvar = caminho & tpEvento & chNFe & nSeqEvento & "-procNFe.xml"
    Else
        localParaSalvar = caminho & tpEvento & chNFe & nSeqEvento & "-procEvenNFe.xml"
    End If

    conteudoSalvar = Replace(xml, "\""", "")

    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudoSalvar
    fsT.SaveToFile localParaSalvar
End Sub

Public Function salvarPDF(pdf As String, caminho As String, chNFe As String, tpEvento As String, nSeqEvento As String) As Boolean
On Error GoTo SAI
    Dim conteudoSalvar  As String
    Dim tpEventoSalvar As String
    Dim localParaSalvar As String

    If (nSeqEvento = "") Then
        localParaSalvar = caminho & chNFe & nSeqEvento & "-procNFe.pdf"
    Else
        localParaSalvar = caminho & chNFe & nSeqEvento & "-procEvenNFe.pdf"
    End If

    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Binary As #fnum
    Put #fnum, 1, Base64Decode(pdf)
    Close fnum
    
    Exit Function
SAI:
    MsgBox (Err.Number & " - " & Err.Description), vbCritical
End Function

Public Function LerDadosJSON(sJsonString As String, key1 As String, key2 As String, key3 As String, Optional key4 As String, Optional key5 As String) As String
On Error GoTo err_handler
    Dim oScriptEngine As ScriptControl
    Set oScriptEngine = New ScriptControl
    oScriptEngine.Language = "JScript"
    Dim objJSON As Object
    Set objJSON = oScriptEngine.Eval("(" + sJsonString + ")")
    If key1 <> "" And key2 <> "" And key3 <> "" And key4 <> "" And key5 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet), key4, VbGet), key5, VbGet)
    ElseIf key1 <> "" And key2 <> "" And key3 <> "" And key4 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet), key4, VbGet)
    ElseIf key1 <> "" And key2 <> "" And key3 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet)
    ElseIf key1 <> "" And key2 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet)
    ElseIf key1 <> "" Then
        LerDadosJSON = VBA.CallByName(objJSON, key1, VbGet)
    End If
Err_Exit:
    Exit Function
err_handler:
    LerDadosJSON = "Error: " & Err.Description
    Resume Err_Exit
End Function

Public Function LerDadosXML(sXml As String, key1 As String, key2 As String) As String
    On Error Resume Next
    LerDadosXML = ""
    
    Set xml = New DOMDocument60
    xml.async = False
    
    If xml.loadXML(sXml) Then
        Set objNodeList = xml.getElementsByTagName(key1 & "//" & key2)
        Set objNode = objNodeList.nextNode
        
        Dim valor As String
        valor = objNode.Text
        
        If Len(Trim(valor)) > 0 Then
            LerDadosXML = valor
        End If
        Else
        MsgBox "Não foi possível ler o conteúdo do XML da NFe especificado para leitura.", vbCritical, "ERRO"
    End If
End Function

Public Sub gravaLinhaLog(conteudoSalvar As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim localParaSalvar As String
    Dim data As String
    
    'Diretório para salvar os logs
    localParaSalvar = App.Path & "\log\"
    
    'Checa se existe o caminho passado para salvar os arquivos
    If Dir(localParaSalvar, vbDirectory) = "" Then
        MkDir (localParaSalvar)
    End If
    
    'Pega data atual
    data = Format(Date, "yyyyMMdd")
    
    'Diretório + nome do arquivo para salvar os logs
    localParaSalvar = App.Path & "\log\" & data & ".txt"
    
    'Pega data e hora atual
    data = DateTime.Now
    
    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Append Shared As #fnum
    Print #fnum, data & " - " & conteudoSalvar
    Close fnum
End Sub
