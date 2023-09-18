# NSNFeAPIClientVB6

Esta página apresenta trechos de códigos de um módulo em VB6 que foi desenvolvido para consumir as funcionalidades da NS NF-e API.

-------

## Primeiros passos:

### Integrando ao sistema:

Para utilizar as funções de comunicação com a API, você precisa realizar os seguintes passos:

1. Extraia o conteúdo da pasta compactada que você baixou;
2. Copie para a pasta da sua aplicação os módulos **NFeAPI.bas** e **base64Convert.bas**, que estão na pasta raiz;
3. Abra o seu projeto e importe a pasta copiada.
4.A aplicação utiliza as bibliotecas **Microsoft Script Control 1.0** e **Active Microsoft XML, v6.0** para realizar a comunicação com a API e fazer a manipulação de dados JSON, respectivamente. Ative as duas referencias em: **Project > References**. 

**Pronto!** Agora, você já pode consumir a NFe em VB6 através do seu sistema.

------

## Emissão Sincrona:

### Realizando uma Emissão:

Para realizar uma emissão completa, você poderá utilizar a função emitirNFeSincrono do módulo NFeAPI. Veja abaixo sobre os parâmetros necessários, e um exemplo de chamada do método.

##### Parâmetros:

**ATENÇÃO:** o **token** também é um parâmetro necessário e você deve primeiramente defini-lo no módulo NFeAPI.bas. Ele é uma constante do módulo. 

Parametros     | Descrição
:-------------:|:-----------
conteudo       | Conteúdo de emissão do documento.
tpConteudo     | Tipo de conteúdo que está sendo enviado. Valores possíveis: json, xml, txt
CNPJ           | CNPJ do emitente do documento.
tpDown         | Tipo de arquivos a serem baixados.Valores possíveis: <ul> <li>**X** - XML</li> <li>**J** - JSON</li> <li>**P** - PDF</li> <li>**XP** - XML e PDF</li> <li>**JP** - JSON e PDF</li> </ul> 
tpAmb          | Ambiente onde foi autorizado o documento.Valores possíveis:<ul> <li>1 - produção</li> <li>2 - homologação</li> </ul>
caminho        | Caminho onde devem ser salvos os documentos baixados.
exibeNaTela    | Se for baixado, exibir o PDF na tela após a autorização.Valores possíveis: <ul> <li>**True** - será exibido</li> <li>**False** - não será exibido</li> </ul> 

##### Exemplo de chamada:

Após ter todos os parâmetros listados acima, você deverá fazer a chamada da função. Veja o código de exemplo abaixo:
           
    Dim retorno As String
    retorno = emitirNFeSincrono(conteudoEnviar, "json", "07364617000135", "XP", "2", "C:\Documentos", True)
    MessageBox(retorno)

A função **emitirNFeSincrono** fará o envio, a consulta e download do documento, utilizando as funções emitirNFe, consultarStatusProcessamento e downloadNFeAndSave, presentes no módulo NFeAPI.bas. Por isso, o retorno será um JSON com os principais campos retornados pelos métodos citados anteriormente. No exemplo abaixo, veja como tratar o retorno da função emitirNFeSincrono:

##### Exemplo de tratamento de retorno:

O JSON retornado pelo método terá os seguintes campos: statusEnvio, statusConsulta, statusDownload, cStat, chNFe, nProt, motivo, nsNRec, erros. Veja o exemplo abaixo:

    {
        "statusEnvio": "200",
        "statusConsulta": "200",
        "statusDownload": "200",
        "cStat": "100",
        "chNFe": "43181007364617000135550000000119741004621864",
        "nProt": "143180007036833",
        "motivo": "Autorizado o uso da NF-e",
        "nsNRec": "313022",
        "erros": ""
    }
      
Confira um código para tratamento do retorno, no qual pegará as informações dispostas no JSON de Retorno disponibilizado:

    Dim retorno As String
    retorno = emitirNFeSincrono(conteudoEnviar, "json", "07364617000135", "XP", "2", "C:\Documentos", True)

    Dim statusEnvio, statusConsulta, statusDownload, cStat, chNFe, nProt, motivo, nsNRec, erros As String

    statusEnvio = LerDadosJSON(retorno, "statusEnvio", "", "")
    statusConsulta = LerDadosJSON(retorno, "statusConsulta", "", "")
    statusDownload = LerDadosJSON(retorno, "statusDownload", "", "")
    cStat = LerDadosJSON(retorno, "cStat", "", "")
    chNFe = LerDadosJSON(retorno, "chNFe", "", "")
    nProt = LerDadosJSON(retorno, "nProt", "", "")
    motivo = LerDadosJSON(retorno, "motivo", "", "")
    nsNRec = LerDadosJSON(retorno, "nsNRec", "", "")
    erros = LerDadosJSON(retorno, "erros", "", "")

    'Testa se houve sucesso na emissão
    If (statusEnvio = 200) Or (statusEnvio = -6) Then
        'Testa se houve sucesso na consulta
        If (statusConsulta = 200) Then
            'Testa se a nota foi autorizada
            If (cStat = 100) Then
                'Aqui dentro você pode realizar procedimentos como desabilitar o botão de emitir, etc
                MsgBox (motivo)

                'Testa se o download teve problemas
                If (statusDownload <> 200) Then
                    MsgBox (motivo)
                End If
            Else
                'Aqui você pode mostrar alguma solução para o parceiro ou exibir opção de editar a nota
                MsgBox (motivo)
            End If
        'Caso tenha dado erro na consulta
        Else
            'Aqui você pode mostrar uma mensagem ao usuário
            MsgBox (motivo + Chr(13) + erros)
        End If
    Else
        'Aqui você pode exibir para o usuário o erro que ocorreu no envio
        MsgBox (motivo + Chr(13) + erros)
    End If  

-----

## Demais Funcionalidades:

No módulo NFeAPI, você pode encontrar também as seguintes funcionalidades:

NOME                     | FINALIDADE             | DOCUMENTAÇÂO CONFLUENCE
:-----------------------:|:----------------------:|:-----------------------
**enviaConteudoParaAPI** |Função genérica que envia um conteúdo para API. Requisições do tipo POST.|
**emitirNFe** | Envia uma NF-e para processamento.|[Emitir NF-e](https://confluence.ns.eti.br/pages/viewpage.action?pageId=13861631#Emiss%C3%A3onaNSNF-eAPI-eAPI-Emiss%C3%A3odeNF-e).
**consultarStatusProcessamento** | Consulta o status de processamento de uma NF-e.| [Status de Processamento da NF-e](https://confluence.ns.eti.br/pages/viewpage.action?pageId=13861631#Emiss%C3%A3onaNSNF-eAPI-StatusdeProcessamentodaNF-e).
**downloadNFe** | Baixa documentos de emissão de uma NF-e autorizada. | [Download da NF-e](https://confluence.ns.eti.br/pages/viewpage.action?pageId=13861631#Emiss%C3%A3onaNSNF-eAPI-DownloaddaNF-e)
**downloadNFeESalvar** | Baixa documentos de emissão de uma NF-e autorizada e salva-os em um diretório. | Por utilizar o método downloadNFe, a documentação é a mesma. 
**downloadEventoNFe** | Baixa documentos de evento de uma NF-e autorizada | [Download de Evento de NF-e](https://confluence.ns.eti.br/display/PUB/Download+de+Evento+na+NS+NF-e+API).
**downloadEventoNFeESalvar** | Baixa documentos de evento de uma NF-e autorizada e salva-os em um diretório. | Por utilizar o método downloadEventoNFe, a documentação é a mesma.
**cancelarNFe** | Realiza o cancelamento de uma NF-e. | [Cancelamento de NF-e](https://confluence.ns.eti.br/display/PUB/Cancelamento+na+NS+NF-e+API).
**corrigirNFe** | Realiza a CC-e de uma NF-e. | [Carta de Correção de NF-e](https://confluence.ns.eti.br/pages/viewpage.action?pageId=13861884).
**consultarCadastroContribuinte** | Consulta o cadastro de um contribuinte. | [Consulta Cadastro de Contribuinte](https://confluence.ns.eti.br/display/PUB/Consulta+Cadastro+de+Contribuinte+na+NS+NF-e+API).
**consultarSituacao** | Consulta a situação de uma NF-e na Sefaz. | [Consulta Situação da NF-e](https://confluence.ns.eti.br/pages/viewpage.action?pageId=13862181).
**enviarEmail** | Envia NF-e por e-mail. (Para enviar mais de um e-mail, separe os endereços por vírgula). | [Envio de NF-e por E-mail](https://confluence.ns.eti.br/display/PUB/Envio+de+NF-e+por+E-mail+na+NS+NF-e+API).
**inutilizar** | Inutiliza numerações de NF-e. | [Inutilização de Numeração](https://confluence.ns.eti.br/pages/viewpage.action?pageId=13862178).
**listarNSNRecs** | Lista os nsNRec vinculados a uma NF-e. | [Lista de NSNRecs vinculados a uma NF-e](https://confluence.ns.eti.br/display/PUB/Lista+de+NSNRecs+vinculados+a+uma+NF-e+na+NS+NF-e+API).
**salvarXML** | Salva um XML em um diretório. | 
**salvarJSON** | Salva um JSON em um diretório. |
**salvarPDF** |	Salva um PDF em um diretório. | 
**LerDadosJSON** | 	Lê o valor de um campo de um JSON. |
**LerDadosXML** | Lê o valor de um campo de um XML. | 
**gravaLinhaLog** | Grava uma linha de texto no arquivo de log. | 



![Ns](https://nstecnologia.com.br/blog/wp-content/uploads/2018/11/ns%C2%B4tecnologia.png) | Obrigado pela atenção!
