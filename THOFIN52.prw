#Include "Protheus.ch"
#Include "FWMVCDef.ch"
#Include "Topconn.ch"
 
/*/{Protheus.doc} THOFIN52
    Rotina responsável por realizar deleção das baixas do cliente por lote
    @type user function
    @author José Vitor Rodrigues
    @since 24/01/2024
    @version 1.0
/*/
User Function THOFIN52()
    // Local cPerg        := "THOFIN52"
    Local aColunas     := {}
    Local nX
    Private oBrowse
    Private cCodCliente, cPrefixo, cNum, cParcDe, cParcAte
    Private aRet       := {}
    Private aArea      := GetArea()
    Private cTempAlias := GetNextAlias()
    Private aCampos    := {}
    Private aCol       := {}
	Private aDados      := {}
    Private Retorno

    IF ! MsgYesNo("AVISO: Essa rotina deve ser executada apenas em situações específicas, caso outras não tenham atendido às necessidades do usuário." + Chr(13) + Chr(10) + "Você deseja prosseguir com a Rotina?" , "Confirma?")
        RETURN    
    ENDIF
    
    IF ! Perguntas()
        RETURN
    ENDIF

    while ((SUBSTR(mv_par04,1,1) != 'E' .AND. SUBSTR(mv_par05,1,1) == 'E'))
        MsgInfo("Coloque o valor de entrada no parametro de!","Atenção") 
        IF ! Perguntas()
            RETURN
        ENDIF     
    end

    //salvando as resposta da pergunta em parametros
    cCodCliente := mv_par01
    cPrefixo := mv_par02
    cNum := mv_par03
    cParcDe := mv_par04
    cParcAte := mv_par05
    
    //Adicionando o campo da tabela temporária
    aadd(aCampos, {"FILIAL"       , "C", 06, 00})
    aadd(aCampos, {"CODPRODUTO"   , "C", 15, 00})
    aadd(aCampos, {"PRODUTO"      , "C", 40, 00})
    aadd(aCampos, {"CODCLIENTE"   , "C", 06, 00})
    aadd(aCampos, {"CLIENTE"      , "C", 50, 00})
    aadd(aCampos, {"PREFIXO"      , "C", 03, 00})
    aadd(aCampos, {"NUM"          , "C", 06, 00})
    aadd(aCampos, {"PARCELA"      , "C", 03, 00})
    aadd(aCampos, {"VALORPARC"    , "C", 15, 00})
    aadd(aCampos, {"ACRESCIMO"    , "C", 15, 00})
    aadd(aCampos, {"VALORTOTAL"   , "C", 15, 00})
    aadd(aCampos, {"BAIXA"        , "C", 10, 00})
    aadd(aCampos, {"MOTIVOBX"     , "C", 03, 00})
    aadd(aCampos, {"SECUR"        , "C", 06, 00})
    aadd(aCampos, {"PARCE"        , "C", 06, 00})
    
    //Criando a tabela temporária
    AliasTMP   := GetNextAlias()
    oTempTable := FWTemporaryTable():New(AliasTMP)
    oTemptable:SetFields(aCampos)
    oTempTable:Create()
    cTitulo := oTempTable::GetTableNameForQuery()

    Processa({|| AdicionarDados()}, "Realizando a consulta...")

    //Adicionando o campo da coluna do brownse
    aadd(aCol, {"FILIAL"       , "C", 01, 00,""})
    aadd(aCol, {"CODPRODUTO"   , "C", 10, 00,""})
    aadd(aCol, {"PRODUTO"      , "C", 12, 00,""})
    aadd(aCol, {"CODCLIENTE"   , "C", 00, 00,""})
    aadd(aCol, {"CLIENTE"      , "C", 16, 00,""})
    aadd(aCol, {"PREF"         , "C", 01, 00,""})
    aadd(aCol, {"TITULO"       , "C", 05, 00,""})
    aadd(aCol, {"PARCELA"      , "C", 01, 00,""})
    aadd(aCol, {"VALOR PARCELA", "C", 03, 00,""})
    aadd(aCol, {"ACRESCIMO"    , "C", 02, 00,""})
    aadd(aCol, {"VALOR TOTAL"  , "C", 04, 00,""})
    aadd(aCol, {"DATA DA BAIXA", "C", 08, 00,""})
    aadd(aCol, {"MOT.BAIXA"    , "C", 00, 00,""})
    aadd(aCol, {"SECUR"        , "C", 03, 00,""})
    aadd(aCol, {"PARCE"        , "C", 03, 00,""})

    For nX := 1 To Len(aCol)    
        AAdd(aColunas,FWBrwColumn():New())
        aColunas[Len(aColunas)]:SetData( &("{||"+aCampos[nX][1]+"}") )
        aColunas[Len(aColunas)]:SetTitle(aCol[nX][1])
        aColunas[Len(aColunas)]:SetSize(aCol[nX][3])
        aColunas[Len(aColunas)]:SetDecimal(aCol[nX][4])              
        aColunas[Len(aColunas)]:SetPicture(aCol[nX][5])              
    Next nX       

    if Retorno  == .T.
        //Criando o FWMarkBrowse
        oBrowse := FWMBrowse():New()
        oBrowse:SetAlias(AliasTMP)               
        oBrowse:SetDescription("Exclusão Baixas por Lote")
        oBrowse:SetColumns(aColunas)
        oBrowse:AddButton("Executar",{|| FWMsgRun(, {|oSay| chamaExclusao(oSay) }, "Processando", "Processando dados")} ,,3,,.F.)
        oBrowse:AddButton("Cancelar", {|| CloseBrowse()},,4,,.F.)
        oBrowse:SetTemporary(.T.)
        oBrowse:DisableDetails()
            
        //Ativando a janela
        oBrowse:Activate() 

        oBrowse:DeActivate()
        FreeObj(oBrowse)
        RestArea( aArea )
    endif

Return 

/*/{Protheus.doc} AdicionarDados()
    Função responsável por adicionar valores a tabela temporária
    @type  Static Function
    @author José Vitor Rodrigues
    @since 24/01/2024
    @version 1.0
/*/
Static Function AdicionarDados()
    Local Parcela := {}
    Local nCont := 0
    cAliasTemp       := " SELECT DISTINCT SE1.E1_FILIAL,SE1.E1_CLIENTE,SE1.E1_PREFIXO,SE1.E1_NUM,"
    cAliasTemp       += "        SE1.E1_PARCELA, SA1.A1_NOME, SB1.B1_DESC,SE1.E1_PRODUTO,"
    cAliasTemp       += "        SE1.E1_VALOR,SE1.E1_ACRESC,(SE1.E1_VALOR + SE1.E1_ACRESC) VALOR,"
    cAliasTemp       += "        SE1.E1_BAIXA, SE5.E5_MOTBX, SE1.E1_ZZSECUR, SE1.E1_ZZPARCE"
    cAliasTemp       += " FROM " + RetSqlName("SE1") + " SE1"
    cAliasTemp       += " INNER JOIN "+ RetSqlName("SA1") +" SA1 ON SE1.E1_CLIENTE = SA1.A1_COD"
    cAliasTemp       += " INNER JOIN "+ RetSqlName("SB1") +" SB1 ON SE1.E1_PRODUTO = SB1.B1_COD"
    cAliasTemp       += " INNER JOIN "+ RetSqlName("SE5") +" SE5 ON SE1.E1_NUM = SE5.E5_NUMERO" 
	cAliasTemp		 +=	"	    AND SE1.E1_PARCELA = SE5.E5_PARCELA AND SE1.E1_FILIAL=SE5.E5_FILIAL" 
    cAliasTemp       += " WHERE SE1.E1_FILIAL = "+xFilial("SE1")
    cAliasTemp       += "       AND SE1.E1_CLIENTE="+cCodCliente
    cAliasTemp       += "       AND SE1.E1_PREFIXO = '"+cPrefixo+"'"
    cAliasTemp       += "       AND SE1.D_E_L_E_T_= ''  "
    cAliasTemp       += "       AND SE1.E1_NUM="+cNum
    if (SUBSTR(cParcDe,1,1) == 'E' .AND. SUBSTR(cParcAte,1,1) != 'E')
        cAliasTemp   += "       AND (SE1.E1_PARCELA BETWEEN '"+cParcDe+"' AND 'E15'"
        cAliasTemp   += "       OR SE1.E1_PARCELA BETWEEN '001' AND '"+cParcAte+"')"
    else
        cAliasTemp   += "       AND SE1.E1_PARCELA BETWEEN '"+cParcDe+"' AND '"+cParcAte+"'"
    endif    
    cAliasTemp       += "       AND SE5.D_E_L_E_T_= ''  "
    cAliasTemp       += " ORDER BY E1_PARCELA  "
    TCQuery cAliasTemp New Alias "QRY_PRO"
    if !(QRY_PRO->(EOF()))
        aadd(aDados,QRY_PRO->E1_FILIAL)
        aadd(aDados,QRY_PRO->E1_PRODUTO)
        aadd(aDados,QRY_PRO->B1_DESC)
        aadd(aDados,QRY_PRO->E1_CLIENTE)
        aadd(aDados,QRY_PRO->A1_NOME)    
        aadd(aDados,QRY_PRO->E5_MOTBX)    

        while !QRY_PRO->(EOF())
            RecLock((AliasTmp),.T.)
                (AliasTmp)->FILIAL     := ALLTRIM(QRY_PRO->E1_FILIAL)
                (AliasTmp)->CODCLIENTE := ALLTRIM(QRY_PRO->E1_CLIENTE)
                (AliasTmp)->CLIENTE    := ALLTRIM(QRY_PRO->A1_NOME)
                (AliasTmp)->CODPRODUTO := ALLTRIM(QRY_PRO->E1_PRODUTO)
                (AliasTmp)->PRODUTO    := ALLTRIM(QRY_PRO->B1_DESC)
                (AliasTmp)->NUM        := ALLTRIM(QRY_PRO->E1_NUM)
                (AliasTmp)->PARCELA    := ALLTRIM(QRY_PRO->E1_PARCELA)
                (AliasTmp)->PREFIXO    := ALLTRIM(QRY_PRO->E1_PREFIXO)
                (AliasTmp)->VALORPARC  := CVALTOCHAR(QRY_PRO->E1_VALOR)
                (AliasTmp)->ACRESCIMO  := CVALTOCHAR(QRY_PRO->E1_ACRESC)
                (AliasTmp)->VALORTOTAL := CVALTOCHAR(QRY_PRO->VALOR)
                (AliasTmp)->MOTIVOBX   := ALLTRIM(QRY_PRO->E5_MOTBX)
                (AliasTmp)->BAIXA      := SUBSTR(QRY_PRO->E1_BAIXA,7,2)+'/'+SUBSTR(QRY_PRO->E1_BAIXA,5,2)+'/'+left(QRY_PRO->E1_BAIXA,4)
                (AliasTmp)->SECUR      := ALLTRIM(QRY_PRO->E1_ZZSECUR)
                (AliasTmp)->PARCE      := ALLTRIM(QRY_PRO->E1_ZZPARCE)
            MsUnLock()
            nCont++
            aadd(Parcela,QRY_PRO->E1_PARCELA) 
            
            QRY_PRO->(DBSKIP())
        end
        IF SUBSTR(Parcela[nCont],1,1) == 'E'
            aadd(aDados,Parcela[nCont-1])
        Else
            aadd(aDados,Parcela[nCont])
            Retorno := .T.
        ENDIF 
    else
        MSGINFO( "Esse cliente não possui nenhuma baixa em aberto", "Informação" )
        Retorno := .F.
    ENDIF
Return 

/*/{Protheus.doc} chamaExclusao
    Função para realizar o chamado para a exlusão das baixas
    @type  Static Function
    @author José Vitor Rodrigues
    @since 25/01/2024
    @version 1.0
/*/
Static Function chamaExclusao(oSay)
    IF MsgYesNo("Você irá excluir baixas no sistema" + Chr(13) + Chr(10)+"LEMBRETE: Avise o setor de CONTABILIDADE por essa exclusão"+Chr(13) + Chr(10) + "Você confirma essa exclusão?" , "Confirma?")
        Processa({|| exclusaoBaixa()}, "Excluíndo as baixas...")
    ENDIF
    
    MsgInfo("Baixas excluídas!","Atenção")
    CloseBrowse()
Return 

/*/{Protheus.doc} exclusaoBaixa
    Função responsável por realizar a exclusao da baixa na SE1 e na SE5
    @type  Static Function
    @author José Vitor Rodrigues
    @since 25/01/2024
    @version 1.0
    @example Exemplos:
        u_zEnvMail("teste@servidor.com.br", "Teste", "Teste TMailMessage - Protheus", , .T.)
/*/
Static Function exclusaoBaixa()
    Local cQueryExclui := ""
    Local cQueryDel    := "" 

    //removendo a baixa do SE1
    cQueryExclui := " UPDATE " + RetSqlName("SE1") + " "
    cQueryExclui += " SET E1_SALDO = E1_VALOR,"
    cQueryExclui += "     E1_SDACRES = E1_ACRESC,"
    cQueryExclui += "     E1_BAIXA = '',"
    cQueryExclui += "     E1_VALLIQ = 0,"
    cQueryExclui += "     E1_STATUS = 'A'"
    cQueryExclui += " WHERE D_E_L_E_T_=''"
    cQueryExclui += "       AND E1_FILIAL ="+xFilial("SE1") 
    cQueryExclui += "       AND E1_CLIENTE ="+cCodCliente
    cQueryExclui += "       AND E1_NUM ="+cNum
    cQueryExclui += "       AND E1_PREFIXO='"+cPrefixo+"'"
    if (SUBSTR(cParcDe,1,1) == 'E' .AND. SUBSTR(cParcAte,1,1) != 'E')
        cQueryExclui   += "       AND (E1_PARCELA BETWEEN '"+cParcDe+"' AND 'E15'"
        cQueryExclui   += "       OR E1_PARCELA BETWEEN '001' AND '"+aDados[7]+"')"
    else
        cQueryExclui   += "       AND E1_PARCELA BETWEEN '"+cParcDe+"' AND '"+aDados[7]+"'"
    endif
    
    TCSqlExec(cQueryExclui)

    //Removendo os dados da SE5
    cQueryDel := " UPDATE " + RetSqlName("SE5") + " "
    cQueryDel += " SET D_E_L_E_T_ = '*'"
    cQueryDel += " WHERE E5_FILIAL ="+xFilial("SE5") 
    cQueryDel += "       AND E5_CLIENTE ="+cCodCliente
    cQueryDel += "       AND E5_PREFIXO ='"+cPrefixo+"'"
    cQueryDel += "       AND E5_NUMERO ="+cNum
    if (SUBSTR(cParcDe,1,1) == 'E' .AND. SUBSTR(cParcAte,1,1) != 'E')
        cQueryDel   += "       AND (E5_PARCELA BETWEEN '"+cParcDe+"' AND 'E15'"
        cQueryDel   += "       OR E5_PARCELA BETWEEN '001' AND '"+aDados[7]+"')"
    else
        cQueryDel   += "       AND E5_PARCELA BETWEEN '"+cParcDe+"' AND '"+aDados[7]+"'"
    endif
    TCSqlExec(cQueryDel)
    EnvioEmail()

    /*UPDATE SE5020 
SET SE5020.D_E_L_E_T_ = TSE5BKP.D_E_L_E_T_
FROM SE5020
INNER JOIN TSE5BKP ON SE5020.R_E_C_N_O_ = TSE5BKP.R_E_C_N_O_
WHERE SE5020.E5_PARCELA = TSE5BKP.E5_PARCELA

UPDATE SE1020 
SET SE1020.E1_SALDO = TSE1BKP.E1_SALDO,
	SE1020.E1_SDACRES = TSE1BKP.E1_SDACRES,
	SE1020.E1_BAIXA = TSE1BKP.E1_BAIXA,
	SE1020.E1_VALLIQ = TSE1BKP.E1_VALLIQ,
	SE1020.E1_STATUS = TSE1BKP.E1_STATUS
FROM SE1020
INNER JOIN TSE1BKP ON SE1020.R_E_C_N_O_ = TSE1BKP.R_E_C_N_O_
WHERE SE1020.E1_PARCELA = TSE1BKP.E1_PARCELA*/
Return 

/*/{Protheus.doc} EnvioEmail
    Função para ser efetuado o disparo do email alertando a contabilidade da alterações da baixa
    @type  Static Function
    @author José Vitor Rodrigues
    @since 26/01/2024
    @version 1.0
/*/
Static Function EnvioEmail()
    Local cAssunto := "AVISO: Exclusão de baixa em lote"
    Local cHtml    := ""
    Local cPara    := GETMV("MV_EMAILP")
    cHtml := '<html>'
    cHtml += '<head>'
    cHtml += '<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">'
    cHtml += '<style type="text/css" style="display:none;"> P {margin-top:0;margin-bottom:0;} </style>'
    cHtml += '</head>'
    cHtml += '<body dir="ltr">'
    cHtml += '<span style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">AVISO: O financeiro realizou a exclusão da baixa em lote do seguinte cliente:</span>'
    cHtml += '<div style="text-align: left; margin: 0px;"><span style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);"><br>'
    cHtml += '</span></div>'
    cHtml += '<div style="text-align: left; margin: 0px;"><span style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Filial: '+aDados[1]+'</span></div>'
    cHtml += '<div style="text-align: left; margin: 0px;"><span style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Código do Produto: '+aDados[2]+'</span></div>'
    cHtml += '<div style="text-align: left; margin: 0px;"><span style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Produto: '+aDados[3]+'</span></div>'
    cHtml += '<div style="text-align: left; margin: 0px;"><span style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Código do cliente: '+aDados[4]+'</span></div>'
    cHtml += '<div class="elementToProof" style="text-align: left; margin: 0px;"><span style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Nome: '+aDados[5]+'</span></div>'
    cHtml += '<div class="elementToProof" style="text-align: left; margin: 0px;"><span style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Parcelas de '+cParcDe+' até '+aDados[7]+'</span></div>'
    cHtml += '<div class="elementToProof" style="text-align: left; margin: 0px;"><span style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Motivo da baixa: '+aDados[6]+'</span></div>'
    cHtml += '<div class="elementToProof" style="text-align: left; margin: 0px;"><span style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Usuário: '+cUserName+'</span></div>'
    cHtml += '<div class="elementToProof" style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">'
    cHtml += '<br>'
    cHtml += '</div>'
    cHtml += '</body>'
    cHtml += '</html>'
	fEnvMail(cAssunto, cHtml, cPara)
return


Static Function fEnvMail(cAssunto, cMensagem, cPara)
	Local aArea        := FWGetArea()
	Local lRet         := .T.
	Local oMsg         := Nil
	Local oSrv         := Nil
	Local nRet         := 0
	Local cFrom        := GETMV('MV_EMAILD') //Alltrim(GetMV("MV_EMAILD"))
	Local cUser        := SubStr(cFrom, 1, At("@", cFrom)-1)
	Local cPass        := GETMV('MV_SENHAEM') //Alltrim(GetMV("MV_SENHAEM"))
	Local cSrvFull     := GETMV('MV_RELSERV') //Alltrim(GetMV("MV_RELSERV"))
	Local cServer      := ""
	Local nPort        := 0
	Local nTimeOut     := 60 //GetMV("MV_RELTIME")
	Local cLog         := ""
	Local lUsaTLS      := .T.
	Default cAssunto   := ""
	Default cMensagem  := ""
	Default cPara  := ""
	
	//Se tiver em branco o destinatario, o assunto ou o corpo do email
	If Empty(cPara) .Or. Empty(cAssunto) .Or. Empty(cMensagem)
		cLog += "001 - Destinatario, Assunto ou Corpo do e-Mail vazio(s)!" + CRLF
		lRet := .F.
	EndIf
	
	//Se tiver ok, continua com a montagem do e-Mail
	If lRet
		cServer      := Iif(':' $ cSrvFull, SubStr(cSrvFull, 1, At(':', cSrvFull)-1), cSrvFull)
		nPort        := Iif(':' $ cSrvFull, Val(SubStr(cSrvFull, At(':', cSrvFull)+1, Len(cSrvFull))), 587)
		
		//Cria a nova mensagem
		oMsg := TMailMessage():New()
		oMsg:Clear()
		
		//Define os atributos da mensagem
		oMsg:cFrom    := cFrom
		oMsg:cTo      := cPara
		oMsg:cSubject := cAssunto
		oMsg:cBody    := cMensagem
		
		//Cria servidor para disparo do e-Mail
		oSrv := tMailManager():New()
		
		//Define se ira utilizar o TLS
		If lUsaTLS
			oSrv:SetUseTLS(.T.)
		EndIf
		
		//Inicializa conexao
		nRet := oSrv:Init("", cServer, cUser, cPass, 0, nPort)
		If nRet != 0
			cLog += "004 - Nao foi possivel inicializar o servidor SMTP: " + oSrv:GetErrorString(nRet) + CRLF
			lRet := .F.
		EndIf
		
		If lRet
			//Define o time out
			nRet := oSrv:SetSMTPTimeout(nTimeOut)
			If nRet != 0
				cLog += "005 - Nao foi possivel definir o TimeOut '"+cValToChar(nTimeOut)+"'" + CRLF
			EndIf
			
			//Conecta no servidor
			nRet := oSrv:SMTPConnect()
			If nRet <> 0
				cLog += "006 - Nao foi possivel conectar no servidor SMTP: " + oSrv:GetErrorString(nRet) + CRLF
				lRet := .F.
			EndIf
			
			If lRet
				//Realiza a autenticacao do usuario e senha
				nRet := oSrv:SmtpAuth(cFrom, cPass)
				If nRet <> 0
					cLog += "007 - Nao foi possivel autenticar no servidor SMTP: " + oSrv:GetErrorString(nRet) + CRLF
					lRet := .F.
				EndIf
				
				If lRet
					//Envia a mensagem
					nRet := oMsg:Send(oSrv)
					If nRet <> 0
						cLog += "008 - Nao foi possivel enviar a mensagem: " + oSrv:GetErrorString(nRet) + CRLF
						lRet := .F.
					EndIf
				EndIf
				
				//Desconecta do servidor
				nRet := oSrv:SMTPDisconnect()
				If nRet <> 0
					cLog += "009 - Nao foi possivel desconectar do servidor SMTP: " + oSrv:GetErrorString(nRet) + CRLF
				EndIf
			EndIf
		EndIf
	EndIf
	
	//Se tiver log de avisos/erros
	If !Empty(cLog)
		
		cLog := "+======================= Envio de eMail =======================+" + CRLF + ;
			"Data  - "+dToC(Date())+ " " + Time() + CRLF + ;
			"Funcao    - " + FunName() + CRLF + ;
			"Para      - " + cPara + CRLF + ;
			"Assunto   - " + cAssunto + CRLF + ;
			"Corpo     - " + CRLF + cMensagem + CRLF + ;
			"Existem mensagens de aviso: "+ CRLF +;
			cLog + CRLF +;
			"+======================= Envio de eMail =======================+"
	EndIf
	
	FWRestArea(aArea)
Return lRet


/*/{Protheus.doc} vF52PRF
Função para realizar a valição do parametro mv_par01 e se
caso não tiver mais que um lote para preencher o mv_par02
e mv_par03 automaticamente
@type user function
@author José Vitor Rodrigues
@since 29/01/2024
@version 1
@param cCodigo, charactere, codigo do cliente

/*/
User Function vTHOFIN52()
    Local aArea := GetArea()
    Local lRet := .T.
    Local cQueryV := ""
    Local cQueryC := ""
    Local n:=0
    Local cPrefixo, cNum
    
    IF( MV_PAR01 != "" .OR. MV_PAR01 != NIL .OR. MV_PAR01 != "      ")
        cQueryV := " SELECT DISTINCT E1_PREFIXO, E1_NUM, E1_CLIENTE, E1_PRODUTO"
        cQueryV += " FROM "+RetSqlName("SE1")
        cQueryV += " WHERE E1_CLIENTE = "+MV_PAR01
        cQueryV += " AND E1_FILIAL = "+xFilial("SE1")

        TCQuery cQueryV New Alias "QRY_V"

        while !QRY_V->(EOF())
            cPrefixo := QRY_V->E1_PREFIXO
            cNum     := QRY_V->E1_NUM
            n++
            QRY_V->(DBSKIP())
        end
        If n == 1
            MV_PAR02 := cPrefixo
            MV_PAR03 := cNum
        else
            if n == 0
                cQueryC := " SELECT A1_NOME"
                cQueryC += " FROM "+RetSqlName("SA1")
                cQueryC += " WHERE A1_COD = '"+MV_PAR01+"'"
                TCQuery cQueryC New Alias "QRY_C"

                if QRY_C->(EOF())
                    alert("Cliente não localizado!"+ Chr(13) + Chr(10) +"Verifique se o código está correto.")
                    lRet := .F.

                else
                    alert("O cliente em questão não tem lotes"+ Chr(13) + Chr(10) +"Verifique se o código está correto.")
                    lRet := .F.
                endif
                QRY_C->(DBCLOSEAREA())
            else
                alert("O cliente possui mais de um lote então preencha os dados seguintes")
            endif
            MV_PAR02 := "   "
            MV_PAR03 := "      "
        endif
        QRY_V->(DBCLOSEAREA())
    Endif
    RestArea(aArea)
Return lRet

Static Function Perguntas()
    Local aPergs    := {}
    Local cCliente  := Space(TamSX3('A1_COD')[01])
    Local cPrefixo  := Space(3)
    Local cTitulo   := Space(6)
    Local cParceDe  := Space(3)
    Local cParceAte := Space(3)

     
    //Adiciona os parâmetros
    aadd(aPergs, {1, "Código do cliente", cCliente , "", "u_vTHOFIN52()", "SA1CLI", ".T.", 30, .T.})
    aadd(aPergs, {1, "Prefixo"          , cPrefixo , "", ".T."          , ""      , ".T.", 30, .T.})
    aadd(aPergs, {1, "Numero do título" , cTitulo  , "", ".T."          , ""      , ".T.", 30, .T.})
    aadd(aPergs, {1, "Parcela de"       , cParceDe , "", ".T."          , ""      , ".T.", 30, .T.})
    aadd(aPergs, {1, "Parcela até"      , cParceAte, "", ".T."          , ""      , ".T.", 30, .T.})
     
    //Se a pergunta foi confirmada
    If ParamBox(aPergs, "Informe os parâmetros", /*aRet*/, /*bOK*/, /*aButtons*/, /*lCentered*/, /*nPosX*/, /*nPosY*/, /*oDlgWizard*/, /*cLoad*/, .F., .F.)
        Return .T.
    EndIf
Return .F.
 
