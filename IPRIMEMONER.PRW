/*/{Protheus.doc} Wanderson Function IMPRIMEMONER
    (long_description)
    @type  Function
    @author Wanderson
    @since 10/04/2022
    @version 1
    /*/
User Function IMPRIMEMONER()
    Local oReport := NIL

	oReport := RptStruc("Vendas Erro")
	oReport:PrintDialog()
Return

Static Function RPTPrint(oReport)
	Local oSection1 := oReport:Section(1)
	Local oSection2 := oReport:Section(2)
	Local cQuery    := ""
	Local cNumSerie := ""
	
    cQuery := " SELECT ZW1_SERIE,ZW1_DOC,ZW1_FILIAL,ZW1_STATUS "   
   	cQuery += " FROM "+RetSQLName("ZW1")+""
   	cQuery += " WHERE D_E_L_E_T_ = '' AND "
   	cQuery += " ZW1_FILIAL = '"+ xFilial("ZW1")+"'  
    cQuery += " ORDER BY ZW1_SERIE,ZW1_DOC"
    cQuery := ChangeQuery( cQuery )
	
		//Verifica se a tabela ja está aberta.
			If Select("TEMP") <> 0
				DbSelectArea("TEMP")
				DbCloseArea()
			EndIf
	
        dbUseArea( .T., "TOPCONN", TcGenQry(,,cQuery), "TEMP", .T., .T. )
			
			DbSelectArea("TEMP")
			TEMP->(dbGoTop())
			
			oReport:SetMeter(TEMP->(LastRec()))
			
		While !EOF()
			If oReport:Cancel()
				Exit
			EndIf
			//Iniciando a primeira seção
			oSection1:Init()
			oReport:IncMeter()
			
			cNumSerie := TEMP->ZW1_SERIE
			IncProc("Imprimindo Serie "+ Alltrim(TEMP->ZW1_SERIE))
			
	//Imprimindo primeira seção:
        oSection1:Cell("ZW1_FILIAL"):SetValue(TEMP->ZW1_FILIAL)
		oSection1:Cell("ZW1_SERIE"):SetValue(TEMP->ZW1_SERIE)			
		oSection1:Printline()
		
		
			//Iniciar a impressão da seção 2
		oSection2:Init()
		
			//verifica se o codigo do cliente é o mesmo, se sim, imprime os dados do pedido
		While TEMP->ZW1_SERIE == cNumSerie
			oReport:IncMeter()
			
		IncProc("Imprimindo Notas..."+ Alltrim(TEMP->ZW1_DOC))
			oSection2:Cell("ZW1_DOC"):SetValue(TEMP->ZW1_DOC)
			
            Do Case
            Case TEMP->ZW1_STATUS == "a"
                oSection2:Cell("ZW1_STATUS"):SetValue("Venda não cancelada no Protheus")
            Case TEMP->ZW1_STATUS == "b"
                oSection2:Cell("ZW1_STATUS"):SetValue("Cancelamento não enviado ao TSS")
            Case TEMP->ZW1_STATUS == "c"
                oSection2:Cell("ZW1_STATUS"):SetValue("Salto de numeração")
            Case TEMP->ZW1_STATUS == "d"
                oSection2:Cell("ZW1_STATUS"):SetValue("Cancelamento sem retorno do TSS")
            Case TEMP->ZW1_STATUS == "f"
                oSection2:Cell("ZW1_STATUS"):SetValue("Cancelamento rejeitado")
            Case TEMP->ZW1_STATUS == "g"
                oSection2:Cell("ZW1_STATUS"):SetValue("Venda com erro")
            EndCase

            
            
			oSection2:Printline()
			
			TEMP->(dbSkip())
			
		EndDo
			oSection2:Finish()
			oReport:ThinLine()
			
			oSection1:Finish()
			
		EndDo
			
Return


Static Function RPTStruc(cNome)
	Local oReport := Nil
	Local oSection1:= Nil
	Local oSection2:= Nil
	
	oReport := TReport():New(cNome,"Relatório Vendas com erro",/**/,{|oRperot| RPTPRINT(oReport)},"Vendas Erro")
	
	oReport:SetPortrait() //Definindo a orientação como retrato
	
	oSection1 := TRSection():New(oReport, "Serie",{"ZW1"}, NIL,.F.,.T.)
	TRCell():New(oSection1,"ZW1_FILIAL"		,"TEMP","FILIAL"  		,"@!",40)
	TRCell():New(oSection1,"ZW1_SERIE"  ,"TEMP","SERIE"	,"@!",200)
	
	oSection2 := TRSection():New(oReport, "Produtos",{"SB1"}, NIL,.F.,.T.)
	TRCell():New(oSection2,"ZW1_DOC"   	,"TEMP","N. Doc."	,"@!",30)
	TRCell():New(oSection2,"ZW1_STATUS"  	,"TEMP","Descrição"	,"@!",100)
	
	oSection1:SetPageBreak(.F.) //Quebra de seção
	

Return (oReport)
