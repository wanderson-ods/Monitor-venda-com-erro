#INCLUDE "PROTHEUS.CH"
#INCLUDE "TCBROWSE.CH"
#INCLUDE "FWMVCDEF.CH"
#INCLUDE "Parmtype.CH"
#INCLUDE "TopConn.CH"

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±³Funcao    ³ MONITORVENDAER ³ Autor ³ Wanderson          ³ Data ³ 22/02/22 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Monitor NFC-e para numerações que estão com erro           ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso		 ³ Automacao Comercial										  ³±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function MONITORVENDAER()
Local cAlias            := "ZW1"
Local cFiltra           := "ZW1_FILIAL == '"+xFilial('ZW1')+"'"
Local aCores            := {}
Private cCadastro       := "Monitor NFC-e"
Private aRotina         := {}
Private cSerieNFCe      := SuperGetMV("MV_SERIEMO",,"Zerado")
Private aIndex          := {}
Private bFiltraBrw      :={|| Filbrowse("ZW1",@aIndex, @cFiltra)}

if cSerieNFCe == "Zerado"
    cSerieNFCe = ""
    if !MsgYesNo("Já foi criado SX3, SX2, SIX da ZW1/ZW2 e parâmetro MV_SERIEMO", "Informações necessarias para uso")
        RETURN
    endif
ENDIF

U_ProcessaMonitor()

Aadd(aRotina, {"Visualizar"     ,"AxVisual"             ,0,1})
Aadd(aRotina, {"Atualizar"      ,"U_ProcessaMonitor"    ,0,2})
Aadd(aRotina, {"Restaura venda" ,"U_VoltaVendaCanc"     ,0,4})
Aadd(aRotina, {"Legenda"        ,"U_LegendaMonitor"     ,0,6})

//Cores legenda
Aadd(aCores, {"ZW1_STATUS == 'a'", "BR_AMARELO"})
Aadd(aCores, {"ZW1_STATUS == 'b'", "BR_AZUL"})
Aadd(aCores, {"ZW1_STATUS == 'c'", "BR_PRETO"})
Aadd(aCores, {"ZW1_STATUS == 'd'", "BR_VIOLETA"})
Aadd(aCores, {"ZW1_STATUS == 'e'", "BR_LARANJA"})
Aadd(aCores, {"ZW1_STATUS == 'f'", "BR_VERMELHO"})
Aadd(aCores, {"ZW1_STATUS == 'g'", "BR_MARROM"})

DbSelectArea(cAlias)
DbSetOrder(1)

Eval(bFiltraBrw)

DBGOTOP()

MBrowse(6,1,22,75,cAlias,,,,,,aCores)

EndFilBrw(cAlias,aIndex)

return 

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±³Funcao    ³ ProcessaMonitor	 ³ Autor ³ Wanderson    ³ Data ³ 22/02/22 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Valida vendas com erro                                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso		 ³ Automacao Comercial										  ³±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function ProcessaMonitor()
Local nX            := 1
Local nY            := 1
Local nZ            := 1
Local aArea         := {}
Local cQuery        := ""
Local aDados        := {}
Local aDocZw1       := {}
Local aSX5Serie     := ""
Local cUltimoNum    := ""
Private aSeries     := {}
Private nPosItem    := 1
Private aVendErro   := {}
Private cPosInicial := 1

cSerieNFCe  := AllTrim(cSerieNFCe)

if cSerieNFCe == ""//Se não tiver parâmetro ou estiver vazio pega todas
    if MsgYesNo("Não foi encontrado pâmetro MV_SERIEMO ou está vazio, deseja validar todas as series existentes na SX5?", "Serie não definida")
        aSer := FwGetSX5("01")
        for nX := 1 to Len(aSer)
            Aadd(aSeries, aSer[nX][3])
        next
    ENDIF
else 
    while nPosItem != 0 //Valida quantas Series tem no parâmetro
        
        nPosItem := AT("/", cSerieNFCe)

        if nPosItem != 0
            Aadd(aSeries, LEFT( cSerieNFCe, nPosItem - 1)) 
            STUFF( cSerieNFCe, 1, nPosItem, "")
        else
            Aadd(aSeries, cSerieNFCe)
        ENDIF
        
    end
ENDIF
aArea := ZW1->(GetArea())

cQuery := " SELECT "
cQuery += " ZW1_DOC AS DOCUMENTO, "
cQuery += " ZW1_SERIE AS SERIE, "
cQuery += " ZW1_PDV AS PDV "
cQuery += " FROM "
cQuery += " "+RetSQLName("ZW1") "
cQuery += " WHERE "
cQuery += " ZW1_FILIAL = '"+ xFilial("ZW1")+"' "
cQuery += " AND D_E_L_E_T_ != '*' "

TCQuery cQuery New Alias "TMPZW1"

while !TMPZW1->(Eof())
    Aadd(aDados, TMPZW1->SERIE+"/"+TMPZW1->DOCUMENTO+"/"+TMPZW1->PDV)
    Aadd(aDocZw1, TMPZW1->DOCUMENTO)
    TMPZW1->(DbSkip())
ENDDO
TMPZW1->(DBCLOSEAREA())
//RestArea(aArea)

if Len(aDados) != 0
    U_ValidaZW1(aDados)
endif



for nX := 1 to Len(aSeries)

    aSX5Serie   := FwGetSX5("01", PADR( aSeries[nX], TamSX3("X5_CHAVE")[1]))
    cUltimoNum  := aSX5Serie[1][4]
    lContinua   := .T.
    

    cPosInicial := U_MonUltiNum(aSeries[nX])
    
    if VAL(cUltimoNum)  - VAL(cPosInicial) > 2500
        if !MsgYesNo("A serie "+aSeries[nX]+" possui muitos registro para processar, então pode ser que demore alguns minutos para terminar o processamento, quer continuar?", "Continuar processando serie")
            lContinua     := .F. 
        ENDIF
    endif

    if lcontinua    
        for nY := VAL(cPosInicial) to VAL(cUltimoNum) - 1
            cDoc := PADL(nY, Len(AllTrim(cUltimoNum)),"0")
            cDoc := PADR( cDoc, TamSX3("L1_DOC")[1])
            lExiste     := .F.

            IncProc("Registro "+cDoc+" Serie "+aSeries[nX])

            for nZ := 1 to LEN(aDados)
                if aDocZw1[nZ] == cDoc
                    lExiste := .T.
                    nZ := LEN(aDados)
                ENDIF
            next

            if !lExiste//Valida se numeração já não existe na ZW1

                cSitua := ""
                cEncontra := ""
                cPdv := ""
                
                aArea := SL1->(GetArea())
                cQuery := " SELECT "
                cQuery += " L1_SITUA AS SITUA, "
                cQuery += " L1_PDV AS PDV "
                cQuery += " FROM "
                cQuery += " "+RetSQLName("SL1") "
                cQuery += " WHERE "
                cQuery += " L1_FILIAL = '"+ xFilial("ZW1")+"' "
                cQuery += " AND D_E_L_E_T_ != '*' AND "
                cQuery += " L1_DOC = '"+cDoc+"' AND "
                cQuery += " L1_SERIE = '"+aSeries[nX]+"' "

                TCQuery cQuery New Alias "TMPL1"

                while !TMPL1->(Eof())

                    cSitua      := TMPL1->SITUA
                    cEncontra   := "SL1"
                    cPdv        := TMPL1->PDV

                    if cSitua == "OK"
                        cSitua = "OK"
                    elseif cSitua == "X0"
                        cSitua = "b"
                    elseif cSitua == "X1"
                        cSitua = "d"
                    elseif cSitua == "X2"
                        cSitua = "a"
                    elseif cSitua == "X3"
                        cSitua = "f"
                    else //cSitua == "ER"
                        cSitua = "g"
                    ENDIF

                    TMPL1->(DbSkip())
                ENDDO

                TMPL1->(DBCLOSEAREA())
                RestArea(aArea)

                if cEncontra == "" //Nao encontrou na SL1 validar se existe na SLX
                    aArea := SF2->(GetArea())
                    cQuery := " SELECT "
                    cQuery += " LX_SITUA AS SITUA, "
                    cQuery += " LX_PDV AS PDV "
                    cQuery += " FROM "
                    cQuery += " "+RetSQLName("SLX") "
                    cQuery += " WHERE "
                    cQuery += " LX_FILIAL = '"+ xFilial("SLX")+"' "
                    cQuery += " AND D_E_L_E_T_ != '*' "
                    cQuery += " AND LX_CUPOM = '"+cDoc+"' "
                    cQuery += " AND LX_SERIE = '"+aSeries[nX]+"' "

                    TCQuery cQuery New Alias "TMPLX"

                    while !TMPLX->(Eof())
                        cSitua      := TMPLX->SITUA
                        cEncontra   := "SLX"
                        cPdv        := TMPLX->PDV

                        if cSitua == "OK"
                            cSitua = "OK"
                        else
                            cSitua = "e"
                        ENDIF
                        

                    TMPLX->(DbSkip())
                    
                    ENDDO
                    TMPLX->(DBCLOSEAREA())
                    RestArea(aArea)
                ENDIF

                if cEncontra == "" //Nao encontrou na SL1 e SLX validar se Está cancelada
                    aArea := SF2->(GetArea())
                    cQuery := " SELECT "
                    cQuery += " F2_DOC AS DOC "
                    cQuery += " FROM "
                    cQuery += " "+RetSQLName("SF2") "
                    cQuery += " WHERE "
                    cQuery += " F2_FILIAL = '"+ xFilial("SF2")+"' "
                    cQuery += " AND D_E_L_E_T_ = '*' "
                    cQuery += " AND F2_DOC = '"+cDoc+"' "
                    cQuery += " AND F2_SERIE = '"+aSeries[nX]+"' "

                    TCQuery cQuery New Alias "TMPF2"

                    while !TMPF2->(Eof())

                        cSitua      := "e"
                        cEncontra   := "SF2"
                        cPdv        := ""

                    TMPF2->(DbSkip())
                    
                    ENDDO
                    TMPF2->(DBCLOSEAREA())
                    RestArea(aArea)
                ENDIF

                if cEncontra == "" //Não existe registro da numeração
                    cSitua      := "c"
                ENDIF

                if cSitua != "OK"
                    aArea := ZW1->(GetArea())
                    DbSelectArea("ZW1")
                        
                    Begin Transaction

                    RecLock("ZW1", .T.)

                        ZW1->ZW1_FILIAL := FwxFilial("ZW1")
                        ZW1->ZW1_DOC    := cDoc
                        ZW1->ZW1_ORIGEM := cEncontra
                        ZW1->ZW1_STATUS := cSitua  
                        ZW1->ZW1_PDV    := ""
                        ZW1->ZW1_DOC    := cDoc
                        ZW1->ZW1_SERIE  := aSeries[nX]
                    
                    ZW1->(MsUnlock())

                    End Transaction
                    RestArea(aArea)
                        
                ENDIF
            endif
        next nY

        aArea := ZW2->(GetArea())
        DbSelectArea("ZW2")
        ZW2->(DbSetOrder(1))
        ZW2->(DbGotop())

        if ZW2->(dbSeek(FwxFilial("ZW2")+aSeries[nX]))
            Begin Transaction

            RecLock("ZW2", .F.)
                Replace ZW2_ULTNUM With cUltimoNum
            ZW2->(MsUnlock())

            End Transaction
            RestArea(aArea)

        else
            Begin Transaction

                RecLock("ZW2", .T.)

                    ZW2->ZW2_FILIAL := FwxFilial("ZW2")
                    ZW2->ZW2_SERIE  := aSeries[nX]
                    ZW2->ZW2_ULTNUM := cUltimoNum
                
                ZW2->(MsUnlock())

            End Transaction
        ENDIF
        RestArea(aArea)
    ENDIF
next nX

return 

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±³Funcao    ³ LegendaMonitor	 ³ Autor ³ Wanderson    ³ Data ³ 22/02/22 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Ajusta legenda                                             ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso		 ³ Automacao Comercial										  ³±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function LegendaMonitor()

Local aLegenda := {}

AADD(aLegenda, {"BR_AMARELO", "Não cancelou fiscal (X2)"})
AADD(aLegenda, {"BR_AZUL", "Não enviou para TSS cancelamento (X0)"})
AADD(aLegenda, {"BR_PRETO", "Numeração não encontrada"})
AADD(aLegenda, {"BR_VIOLETA", "Cancelamento sem retorno do TSS (X1)"})
AADD(aLegenda, {"BR_LARANJA", "Erro na inutilização"})
AADD(aLegenda, {"BR_VERMELHO", "Cancelamento Rejeitado"})
AADD(aLegenda, {"BR_MARROM", "Venda com erro de processamento"})

BrwLegenda(cCadastro,"Legenda", aLegenda)

return 

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±³Funcao    ³ VoltaVendaCanc	 ³ Autor ³ Wanderson    ³ Data ³ 22/02/22      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Restaura venda com erro, rejeirada e rejeição de cancelamento³   ±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso		 ³ Automacao Comercial										       ³±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function VoltaVendaCanc()

Local cStatus   := ZW1->ZW1_STATUS
Local cSerie    := ZW1->ZW1_SERIE
Local cPDV      := ZW1->ZW1_PDV
Local cDoc      := ZW1->ZW1_DOC
Local aNFCeID	:= {}		//vetor de NFEID (L1_SERIE+L1_DOC) que serao consultados no TSS
Local aNFCeDados:= {}	    //vetor retornado pelo TSS com o status de cada nota que foi consultado

if cStatus == "g" .or. cStatus =="e" .or. cStatus =="f".or. cStatus =="b".or. cStatus =="d"

    if cStatus == "g" .or. cStatus =="f"
        aArea := SL1->(GetArea())
        DbSelectArea("SL1")
        SL1->(DbSetOrder(2))
        SL1->(DbGotop())

        if SL1->(dbSeek(FwxFilial("SL1")+cSerie+cDoc+cPdv))
            Begin Transaction
            if cStatus == "g"
                cMuda := "RX"
            else
                cMuda := "OK"
            endif
            

            RecLock("SL1", .F.)
                Replace L1_SITUA With cMuda
            SL1->(MsUnlock())

            End Transaction
        ENDIF
        RestArea(aArea)

    elseif cStatus == "b" .or. cStatus =="d" //Valida se notas com X0 ou X1 estão canceladas na Sefaz para cancelar no Protheus
        aNFCeID := { cSerie + cDoc}
        //Obtemos os status da nota no TSS
        aNFCeDados := LjNFCeGtID(aNFCeID, "65")

        If Len(aNFCeDados) > 0
            if aNFCeDados[1][4] == "7" .AND. aNFCeDados[1][5] == "2"

                cMuda := "X2"

                aArea := SL1->(GetArea())
                DbSelectArea("SL1")
                SL1->(DbSetOrder(2))
                SL1->(DbGotop())
                if SL1->(dbSeek(FwxFilial("SL1")+cSerie+cDoc+cPdv))
                    Begin Transaction

                    RecLock("SL1", .F.)
                        Replace L1_SITUA With cMuda
                    SL1->(MsUnlock())

                    End Transaction
                ENDIF
                RestArea(aArea)
            ENDIF
        ENDIF

    else
        aArea := SLX->(GetArea())
        DbSelectArea("SLX")
        SLX->(DbSetOrder(1))
        SLX->(DbGotop())

        while !SLX->(EOF())
            if(SLX->LX_CUPOM == cDoc .and. SLX->LX_SERIE == cSerie)
                Begin Transaction
                cMuda := "X0"

                RecLock("SLX", .F.)
                    Replace LX_SITUA With cMuda
                SLX->(MsUnlock())

                End Transaction
                RestArea(aArea)

                Exit
            ENDIF
            dbSkip()
        end

        MsgInfo("Venda restaurada com sucesso!", "Restaura Venda")
    endif
else
    MsgAlert("Esta venda não pode ser processada somente as pendentes de inutilizar e as com erro do processamento.", "Não processada")
endif

return 

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±³Funcao    ³ ValidaZW1	 ³ Autor ³ Wanderson    ³ Data ³ 22/02/22     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Valida se mudou status de vendas com erro                  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso		 ³ Automacao Comercial										  ³±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function ValidaZW1(aZW1)
Local nX            := 1
Local cNumero       := ""
Local cSerie        := ""
Local cPDV          := ""
Local lContinua     := .T. 
Local nQuantidade   := Len(aZW1)  

if nQuantidade > 2500
    if !MsgYesNo("Sua tabela ZW1 vai ser analisada para ver se houve mudança em algum status das vendas com erro, mas nela possui uma quantidade grande de registros por este motivo deve demorar, deseja continuar a validação na ZW1?", "Continuar processando ZW1")
        lContinua     := .F. 
    ENDIF
endif

if lContinua
    for nX := 1 to nQuantidade
        cEncontra   := ""
        cSitua      := "" //Qual situação atual caso tenho corrigido deixar como OK para deletar da ZW1
        

        nDivisao    := AT("/", aZW1[nX])
        cSerie      := LEFT( aZW1[nX], nDivisao - 1)
        STUFF( aZW1[nX], 1, nDivisao, "")
        cNumero     := LEFT( aZW1[nX], nDivisao - 1)
        cPDV        := RIGHT( aZW1[nX], nDivisao + 1)

        IncProc("Registro "+cNumero+" Serie "+cSerie)

        aArea := SL1->(GetArea())
        DbSelectArea("SL1")
        SL1->(DbSetOrder(2))
        SL1->(DbGotop())

        if SL1->(dbSeek(FwxFilial("SL1")+cSerie+cNumero+cPDV))
            cSitua := SL1->L1_SITUA
            cEncontra := "SL1"

            if cSitua == "OK"
                cSitua = "OK"
            elseif cSitua == "X0"
                cSitua = "b"
            elseif cSitua == "X1"
                cSitua = "d"
            elseif cSitua == "X2"
                cSitua = "a"
            elseif cSitua == "X3"
                cSitua = "f"
            elseif cSitua == "ER"
                cSitua = "g"
            ENDIF
        ENDIF


        RestArea(aArea)

        if cEncontra == "" //Nao encontrou na SL1 validar se existe na SLX
            aArea := SLX->(GetArea())
            DbSelectArea("SLX")
            SLX->(DbSetOrder(1))
            SLX->(DbGotop())

            while !SLX->(EOF())
                if(SLX->LX_CUPOM == cNumero .and. SLX->LX_SERIE == cSerie)
                    cSitua      := SLX->LX_SITUA
                    cEncontra   := "SLX"

                    if cSitua == "X0" .or. cSitua == "X1" .or. cSitua == "X2"
                        cSitua = "e"
                    else
                        cSitua = "OK"
                    ENDIF
                    Exit
                ENDIF
                DBSkip()
            end
            RestArea(aArea)
        ENDIF

        if cEncontra == "" //Nao encontrou na SL1 e SLX validar se Está cancelada
            aArea := SF2->(GetArea())
            DbSelectArea("SF2")
            SF2->(DbSetOrder(1))
            SF2->(DbGotop())

            while !SF2->(EOF())
                if(SF2->F2_DOC == cNumero .and. SF2->F2_SERIE == cSerie)
                    cSitua      := SF2->D_E_L_E_T_
                    cEncontra   := "SF2"

                    if cSitua == "*"
                        cSitua = "OK"
                    else
                        cSitua = "e"
                    ENDIF

                    Exit
                ENDIF
                dbSkip()
            end
            RestArea(aArea)
        ENDIF

        if cEncontra == "" //Não existe registro da numeração
            cSitua      := "c"
        ENDIF


        aArea := ZW1->(GetArea())
        DbSelectArea("ZW1")
        ZW1->(DbSetOrder(1))
        ZW1->(DbGotop())

        if ZW1->(dbSeek(FwxFilial("ZW1")+cNumero+cSerie))

            if cSitua == "OK"
                DBDELETE()
            else
                Begin Transaction

                RecLock("ZW1", .F.)
                    Replace ZW1_STATUS With cSitua
                ZW1->(MsUnlock())

                End Transaction
                RestArea(aArea)
            ENDIF
        ENDIF
    next nX
ENDIF

return 

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±³Funcao    ³ MonUltiNum    	 ³ Autor ³ Wanderson    ³ Data ³ 22/02/22 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Pega o ultimo numero na ZW2                                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso		 ³ Automacao Comercial										  ³±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function MonUltiNum(cSer)
Local cNum := "1"

aArea := ZW2->(GetArea())
        DbSelectArea("ZW2")
        ZW2->(DbSetOrder(1))
        ZW2->(DbGotop())

        if SL1->(dbSeek(FwxFilial("ZW2")+cSer))
            cNum      := AllTrim(ZW2->ZW2_ULTNUM)
        ENDIF
        RestArea(aArea)

return cNum
