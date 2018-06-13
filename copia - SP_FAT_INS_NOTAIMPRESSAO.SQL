USE [Corporate]
GO
/****** Object:  StoredProcedure [dbo].[SP_FAT_INS_NOTAIMPRESSAO]    Script Date: 14/05/2018 14:47:52 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[SP_FAT_INS_NOTAIMPRESSAO] 
											  @COD_FILIAL    INT,
                                              @COD_REGIONAL  INT,
                                              @COD_NF        INT,
                                              @COD_SERIENOTA CHAR(2)
AS            
          
             
/*            
CRIADA EM                    : 21/09/2011            
AUTOR                        : PEDRO HENRIQUE RELVAS            
SISTEMA QUE UTILIZA          : FATWIN            
DESCRIÇÃO                    : GRAVA REGISTRO DE CRIAÇÃO DA NF PARA CONGELAMENTO DOS DADOS            
NOME DA PROCEDURE QUE UTILIZA: SP_FAT_INS_NOTAIMPRESSAO            
             
ÚLTIMA ALTERAÇÃO EM          : 16/11/2011            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : ALTERAÇÃO DE CAMPOS EM TABELA DO SICOM E FOI NECESSÁRIO O AJUSTE NESSA PROC            
             
ÚLTIMA ALTERAÇÃO EM          : 17/12/2011            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : NOVO CAMPO CRIADO E PREENCHIMENTO DO MESMO "DES_MSGALIQUOTA"            
             
ÚLTIMA ALTERAÇÃO EM          : 17/12/2011            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : FORMATAÇÃO DOS VALORES DO CAMPO "DES_MSGALIQUOTA"            
             
ÚLTIMA ALTERAÇÃO EM          : 17/12/2011            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : SDM-9554 - FORMATAÇÃO DA DESCRIÇÃO E VALOR DO PRODUTO            
             
ÚLTIMA ALTERAÇÃO EM          : 19/12/2011            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : SDM-9554 - ALTERAÇÃO NO INSERT NA TABELA #DADOS NO FINAL DA PROCEDURE            
            
ÚLTIMA ALTERAÇÃO EM          : 22/12/2011            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : SDM-9554 - ALIMENTAÇÃO DO CAMPO "VAL_INSS" E AJUSTADO CALCULO DE PIS, COFINS E CSLL            
            
ÚLTIMA ALTERAÇÃO EM          : 23/12/2011            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : SDM-9554 - CRIADA CRITICA PARA A MENSAGEM DO ISS E DAS ALIQUOTAS            
            
ÚLTIMA ALTERAÇÃO EM          : 26/12/2011            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : SDM-9554 - AJUSTE NA EXIBIÇÃO DA MENSAGEM DE "ISS RECOLHIDO PELO PRESTADOR"            
            
ÚLTIMA ALTERAÇÃO EM          : 05/01/2012            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : SDM-9554 - AJUSTE NO CRITÉRIO DO PREENCHIMENTO DO CFOP(COD_NATUREZA)            
            
ÚLTIMA ALTERAÇÃO EM          : 20/01/2012            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : SDM-9554 - DEVIDO A MUDANÇA NA LEGISLAÇÃO, A CRITICA DO VALOR DO INSS FOI ALTERADA DE ">=29" PARA ">=10"            
            
ÚLTIMA ALTERAÇÃO EM          : 13/02/2012            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : SDM-9554 - ADICIONADA CRITICA PARA O CAMPO VAL_INSS            
            
ÚLTIMA ALTERAÇÃO EM          : 28/02/2012            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS - PLUS-IT            
MOTIVO DA ALTERAÇÃO          : SDM-9554 - CORREÇÃO NA ALIMENTAÇÃO DO SETOR            
          
ÚLTIMA ALTERAÇÃO EM          : 01/03/2012            
AUTOR DA ALTERAÇÃO           : RICARDO FIRMINO DOS SANTOS           
MOTIVO DA ALTERAÇÃO          : CORREÇÃO NA ATULIZAÇÃO DA GRAFICA            
          
ÚLTIMA ALTERAÇÃO EM          : 07/03/2012            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS          
MOTIVO DA ALTERAÇÃO          : CORREÇÃO NO CAMPO "CGC_FILIAL" E ALTERADO O "INNER JOIN" COM A TABELA FAT_MOVIMENTO E A SIC_PRODUTO PARA "LEFT JOIN"          
          
ÚLTIMA ALTERAÇÃO EM          : 16/03/2012            
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS (PLUS-IT)          
MOTIVO DA ALTERAÇÃO          : ADIÇÃO DO UPDATE DO CAMPO "DES_MSGNF" NA DA TABELA FAT_NOTA, COM O MESMO DADO DA TABELA FAT_NOTAIMPRESSAO          
          
ÚLTIMA ALTERAÇÃO EM          : 05/04/2012            
AUTOR DA ALTERAÇÃO           : HERBERT BRUCHMANN JUNIOR (COMMITMENT ALL EASY)          
MOTIVO DA ALTERAÇÃO          : AJUSTES PARA TRAZER OS CAMPOS REFERENTES AOS CANCELAMENTOS DE NFS (DT_CANCELAMENTO, COD_FUNCCANCELAMENTO, DES_CANCELAMENTO)          
          
ÚLTIMA ALTERAÇÃO EM          : 05/04/2012            
AUTOR DA ALTERAÇÃO           : ANDRÉ SBAMPATO          
MOTIVO DA ALTERAÇÃO          : ADIÇÃO DE VALOR DE INSS NA MENSAGEM          
          
ÚLTIMA ALTERAÇÃO EM          : 09/04/2012          
AUTOR DA ALTERAÇÃO           : PEDRO HENRIQUE RELVAS (PLUS-IT)          
MOTIVO DA ALTERAÇÃO          : PREENCHIMENTO DOS CAMPOS "DESTACAR_ICMS" E "DESTACAR_ISS"          
          
ÚLTIMA ALTERAÇÃO EM          : 28/04/2012          
AUTOR DA ALTERAÇÃO           : ANDRÉ SBAMPATO          
MOTIVO DA ALTERAÇÃO          : ADIÇÃO DE INFORMAÇÕES NO CORPO DA NOTA PARA NOTAS ELETRÔNICAS          
          
ÚLTIMA ALTERAÇÃO EM          : 19/06/2012          
AUTOR DA ALTERAÇÃO           : ANDRÉ SBAMPATO          
MOTIVO DA ALTERAÇÃO          : ACERTO DE MENSAGEM DE SERVIÇO          
          
ÚLTIMA ALTERAÇÃO EM          : 03/09/2012          
AUTOR DA ALTERAÇÃO           : ANDRÉ SBAMPATO          
MOTIVO DA ALTERAÇÃO          : INCLUSÃO DE CAMPOS DE NFS-e          
          
ÚLTIMA ALTERAÇÃO EM          : 18/10/2012          
AUTOR DA ALTERAÇÃO           : T4 - SILVIO           
MOTIVO DA ALTERAÇÃO          : SDM-11377: INCLUSÃO TRATAMENTO FAMÍLIA 12 - LOCAÇÃO COMPUSAFE          
        
ÚLTIMA ALTERAÇÃO EM          : 15/02/2013          
AUTOR DA ALTERAÇÃO           : ANDRÉ SBAMPATO          
MOTIVO DA ALTERAÇÃO          : INCLUSÃO DE CÓDIGO FISCAL ISS PARA DESEMBARAÇO ADUANEIRO E AGENCIAMENTO DE CARGA         
          
ÚLTIMA ALTERAÇÃO EM          : 28/02/2013          
AUTOR DA ALTERAÇÃO           : ANDRÉ SBAMPATO          
MOTIVO DA ALTERAÇÃO          : GERAÇÃO DE MENSAGEM ESPECÍFICA DE RECIBO DA BVA        
          
ÚLTIMA ALTERAÇÃO EM          : 11/04/2014            
AUTOR DA ALTERAÇÃO           : ANDRÉ SBAMPATO            
MOTIVO DA ALTERAÇÃO          : TUNNING PARA ACELERAR A GERAÇÃO DA NOTA          
        
ÚLTIMA ALTERAÇÃO EM          : 27/03/2014          
AUTOR DA ALTERAÇÃO           : NELSON C. PRADO - TALENTFOUR          
MOTIVO DA ALTERAÇÃO          : AJUSTE TEXTO LEGAL CONF. SDM 11673           
        
ÚLTIMA ALTERAÇÃO EM          : 08/04/2014          
AUTOR DA ALTERAÇÃO           : NELSON C. PRADO - TALENTFOUR          
MOTIVO DA ALTERAÇÃO          : Inclusão dos campos para validação da isenção de ISS,ICMS. PIS, CONFINS, CSLL, IRRF     
SDM        : 11673          
  
ÚLTIMA ALTERAÇÃO EM          : 05/05/2014   
AUTOR DA ALTERAÇÃO           : VALDERIO JUNIOR    
MOTIVO DA ALTERAÇÃO          : SDM-12332: INCLUSÃO TRATAMENTO FAMÍLIA 13 - CONSULTORIA EM SEGURANÇA  
  
ÚLTIMA ALTERAÇÃO EM          : 26/01/2015  
AUTOR DA ALTERAÇÃO           : EDUARDO BOQUETTI (PLUS-IT)  
MOTIVO DA ALTERAÇÃO          : SDM 12699 UC01 - Correcao de CFOP e quebra por estado(UF) de arrecadação icms  
  
ÚLTIMA ALTERAÇÃO EM          : 10/07/2015  
AUTOR DA ALTERAÇÃO           : ANDRÉ SBAMPATO  
MOTIVO DA ALTERAÇÃO          : SDM 12908: Nova regra federal de PIS\COFINS\CSLL  
  
ÚLTIMA ALTERAÇÃO EM          : 12/10/2015    
AUTOR DA ALTERAÇÃO           : ANDRÉ SBAMPATO  
MOTIVO DA ALTERAÇÃO          : 12981: Opção de GTVs refaturadas  
  
ÚLTIMA ALTERAÇÃO EM          : 17/06/2016  
AUTOR DA ALTERAÇÃO           : Gabriel Godoi   
MOTIVO DA ALTERAÇÃO          : Incluída lógica do CFOP para notas de BSB  
SDM        : 13121    
  
ÚLTIMA ALTERAÇÃO EM          : 29/06/2016  
AUTOR DA ALTERAÇÃO           : KAIO FELIPE  
MOTIVO DA ALTERAÇÃO          : Correção para coletar apenas registros ativos na SIC_TRIBISSNACIONAL  
  
ÚLTIMA ALTERAÇÃO EM          : 03/08/2016  
AUTOR DA ALTERAÇÃO           : Gustavo Gobbi (Plus-IT)  
MOTIVO DA ALTERAÇÃO          : Inclusao do Parametro de Cobrança, UF_destino e metodo do Oracle  
SDM        : 13333  
  
ÚLTIMA ALTERAÇÃO EM          : 26/08/2016  
AUTOR DA ALTERAÇÃO           : ANDRÉ SBAMPATO  
MOTIVO DA ALTERAÇÃO          : Acerto de alíquota de IRRF  
  
ÚLTIMA ALTERAÇÃO EM          : 02/09/2016  
AUTOR DA ALTERAÇÃO           : Gustavo Gobbi (Plus-IT)  
MOTIVO DA ALTERAÇÃO          : NAO INCLUIR DES_METODO PARA COBRANCA DIFERENTE DE BOLETO  
SDM        : 13333  
  
ÚLTIMA ALTERAÇÃO EM          : 30/11/2016  
AUTOR DA ALTERAÇÃO           : Liryanne Nogueira (Plus-IT)  
MOTIVO DA ALTERAÇÃO          : Adicionado CREATE para a tabela #FAT_MOVIMENTO; adicionado os campos no SELECT da FAT_MOVIMENTO que insere dados na #FAT_MOVIMENTO.  
  
ÚLTIMA ALTERAÇÃO EM          : 29/12/2016  
AUTOR DA ALTERAÇÃO           : ANDRÉ SBAMPATO  
MOTIVO DA ALTERAÇÃO          : Acerto emergencial em produção, chamado: INC1734151  
  
ÚLTIMA ALTERAÇÃO EM          : 05/01/2017  
AUTOR DA ALTERAÇÃO           : KAIO FELIPE  
MOTIVO DA ALTERAÇÃO          : Acerto emergencial em produção, chamado: INC1734151  
  
ÚLTIMA ALTERAÇÃO EM          : 22/06/2017  
AUTOR DA ALTERAÇÃO           : FERNANDO FONTE (PLUS-IT) / JOSÉ GABRIEL (BRINKS)  
MOTIVO DA ALTERAÇÃO          : SDM 13206 - INCLUSÃO DO CAMPO: Des_CidOrigemCTe, Cod_IBGEOrigemCTe, Des_CidDestinoCTe, Cod_IBGEDestinoCTe  

ÚLTIMA ALTERAÇÃO EM          : 04/05/2018  
AUTOR DA ALTERAÇÃO           : Ricardo Gomes (PLUS-IT) 
MOTIVO DA ALTERAÇÃO          : WKF 255294 - Gerar descrição de impostos IRRF, PCLL e INSS buscando a descrição das mensagens em tabelas
  
Última alteração em          : 19/05/2018
Autor da alteração           : Carlos Alexandre Negri (Interaktiv)
Motivo da alteração          : Alteração do tamanho do campo [Des_NumBoleto] da tabela [#FAT_NOTAIMPRESSAO] de 12 para 13 para o boleto Santander 
SDM                          : 172738

EXEMPLO       :EXEC SP_FAT_INS_NOTAIMPRESSAO 83,1,623,'1'          
         
*/

     SET NOCOUNT ON;
     SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;
     DECLARE @PERCENTUAL DECIMAL, @LOTEFORMULARIO VARCHAR(250), @REGIMEESPECIALESTADO VARCHAR(200), @NOMEDAGRAFICA VARCHAR(250), @REGIMEESPECIALMUNICIPIO VARCHAR(200), @COD_ISS CHAR(10), @CONTADOR INT, @COD_SERVICO VARCHAR(5);  
  
--CONFIGURAÇÕES E VARIÁVEIS PARA REGRAS DE IMPOSTOS FEDERAIS  
     DECLARE @BO_CALCULA_IMPOST_FEDERAIS BIT, @BO_ISENTO_IMPOST_FEDERAIS BIT, @FED_VAL_MINIMOIMPOSTO FLOAT, @FED_VAL_MINIMONOTA FLOAT, @DES_MSGSemRetencao VARCHAR(110), @DES_MSGSemRetencaoAlt VARCHAR(110), @DES_MSGComRetencao VARCHAR(110), @DES_TotAliqFederal VARCHAR(4), @DES_MSGSemRetencaoFinal VARCHAR(110);  

--<WKF 255294 - Ricardo Gomes - Plus It>
     DECLARE
      @MSG_IRRF VARCHAR(250),
      @MSG_PCC VARCHAR(250),
      @ALIQ_IRRFTRIMEMP DECIMAL(5,2),
      @MSG_INSS VARCHAR(250),
      @MSG_INSSNRETIDO VARCHAR(250),
      @VALORMINIMPOSTO FLOAT,
      @VAL_PCC_MSG VARCHAR(20),
      @IMPOSTO_PCC_MSG VARCHAR(10)

--</WKF 255294 - Ricardo Gomes - Plus It>

     SET @PERCENTUAL = 0;            
            
--VERIFICA SE A NF EXISTE NA TABELA DE IMPRESSAO, SE SIM EXCLUI O REGISTRO            
     IF EXISTS
(
    SELECT *
    FROM FAT_NOTAIMPRESSAO
    WHERE COD_FILIAL = @COD_FILIAL
          AND COD_REGIONAL = @COD_REGIONAL
          AND COD_NF = @COD_NF
          AND COD_SERIENOTA = @COD_SERIENOTA
)
         BEGIN
             DELETE FROM FAT_NOTAIMPRESSAO
             WHERE COD_FILIAL = @COD_FILIAL
                   AND COD_REGIONAL = @COD_REGIONAL
                   AND COD_NF = @COD_NF
                   AND COD_SERIENOTA = @COD_SERIENOTA;
         END;            
         
  
--*********** CARREGA CONFIGURAÇÕES  

     SELECT @FED_VAL_MINIMOIMPOSTO = VAL_MINIMOIMPOSTO,
            @FED_VAL_MINIMONOTA = VAL_MINIMONOTA,
            @DES_MSGSemRetencao = DES_MSGSemRetencao,
            @DES_MSGSemRetencaoAlt = DES_MSGSemRetencaoAlt,
            @DES_MSGComRetencao = DES_MSGCOMRETENCAO
     FROM FAT_ConfigImpostosFederais
     WHERE Flg_situacao = 1;  
             
--CREATE TABLES--------------------------------------------------------------------------------------            
      
--TEMPORÁRIA DA TABELA ESPELHO          
     CREATE TABLE #FAT_NOTAIMPRESSAO
([COD_FILIAL]                   [INT] NOT NULL,
 [COD_REGIONAL]                 [INT] NOT NULL,
 [COD_NF]                       [INT] NOT NULL,
 [COD_SERIENOTA]                [CHAR](2) NOT NULL,
 [BAS_ICMS]                     [FLOAT] NULL,
 [BAS_ISS]                      [FLOAT] NULL,
 [CEP_CLIENTE]                  [CHAR](10) NULL,
 [CEP_COBRANCA]                 [CHAR](10) NULL,
 [CEP_FILIAL]                   [CHAR](10) NULL,
 [CGC_CLIENTE]                  [VARCHAR](30) NULL,
 [CGC_CLIENTE2]                 [VARCHAR](30) NULL,
 [CGC_FILIAL]                   [VARCHAR](50) NULL,
 [CGC_ORIGEM]                   [VARCHAR](15) NULL,
 [CGC_SUBCLIENTENF]             [VARCHAR](20) NULL,
 [CGC_SUBCLIENTESACADO]         [VARCHAR](20) NULL,
 [CIDADE_CLIENTE]               [VARCHAR](30) NULL,
 [COD_BARRABOLETOBANC]          [VARCHAR](44) NULL,
 [COD_CGC]                      [VARCHAR](20) NULL,
 [COD_CLIENTETRIBUTACAO]        [INT] NULL,
 [COD_EMPRESA]                  [INT] NULL,
 [COD_FAMILIAISS]               [SMALLINT] NULL,
 [COD_FISCAL]                   [INT] NULL,
 [COD_FISCAL_ISS]               [NUMERIC](5, 0) NULL,
 [COD_FUNCCANCELAMENTO]         [INT] NULL,
 [COD_FUNCEMISSAO]              [INT] NULL,
 [COD_FUNCNFSUBS]               [INT] NULL,
 [COD_INTERESTADUAL]            [INT] NULL,
 [COD_INTERMUNICIPAL]           [INT] NULL,
 [COD_LINHADIGBOLETOBANC]       [VARCHAR](47) NULL,
 [COD_NATUREZA]                 [VARCHAR](5) NULL,
 [COD_NFSUBSTITUTA]             [INT] NULL,
 [COD_RAMO]                     [INT] NULL,
 [COD_RECIBO]                   [INT] NULL,
 [COD_TRIBUTACAO]               [INT] NULL,
 [CON_CORRENTESANTANDER]        [VARCHAR](200) NULL,
 [DES_BAIRROCOBRANCA]           [VARCHAR](60) NULL,
 [DES_CANCELAMENTO]             [VARCHAR](200) NULL,
 [DES_CIDADE]                   [VARCHAR](200) NULL,
 [DES_CIDADECOBRANCA]           [VARCHAR](30) NULL,
 [DES_ENDCOBRANCA]              [VARCHAR](90) NULL,
 [DES_IRRF]                     [VARCHAR](200) NULL,
 [DES_JUSTFORASEQ]              [VARCHAR](100) NULL,
 [DES_MENSAGEMSUBS_ICMS]        [VARCHAR](200) NULL,
 [DES_MENSAGEMSUBS_ISS]         [VARCHAR](200) NULL,
 [DES_MSGBOLETO]                [VARCHAR](30) NULL,
 [DES_MSGNF]                    [VARCHAR](300) NULL,
 [DES_NUMBOLETO]                [VARCHAR](13) NULL,
 [DES_PRODUTO_ICMS1]            [VARCHAR](100) NULL,
 [DES_PRODUTO_ICMS2]            [VARCHAR](100) NULL,
 [DES_PRODUTO_ICMS3]            [VARCHAR](100) NULL,
 [DES_PRODUTO_ICMS4]            [VARCHAR](100) NULL,
 [DES_PRODUTO_ICMS5]            [VARCHAR](100) NULL,
 [DES_PRODUTO_ISS1]             [VARCHAR](100) NULL,
 [DES_PRODUTO_ISS2]             [VARCHAR](100) NULL,
 [DES_PRODUTO_ISS3]             [VARCHAR](100) NULL,
 [DES_PRODUTO_ISS4]             [VARCHAR](100) NULL,
 [DES_PRODUTO_ISS5]             [VARCHAR](100) NULL,
 [DES_PRODUTO_ISS6]             [VARCHAR](100) NULL,
 [DES_QTD_ICMS]                 [VARCHAR](100) NULL,
 [DES_QTD_ISS]                  [VARCHAR](100) NULL,
 [DES_SETOREXTERNO]             [VARCHAR](20) NULL,
 [DES_SUBCLIENTE]               [VARCHAR](200) NULL,
 [DES_SUBSTRIBICMS]             [VARCHAR](200) NULL,
 [DES_TPSUBCLIENTE]             [VARCHAR](30) NULL,
 [DES_UNID_ICMS]                [VARCHAR](100) NULL,
 [DES_UNID_ISS]                 [VARCHAR](100) NULL,
 [DES_VALOR_ICMS]               [VARCHAR](100) NULL,
 [DES_VALOR_ISS]                [VARCHAR](100) NULL,
 [DES_VALOREXTENSO]             [VARCHAR](120) NULL,
 [DESTACAR_ICMS]                [CHAR](1) NULL,
 [DESTACAR_ISS]                 [CHAR](1) NULL,
 [DT_CANCELAMENTO]              [SMALLDATETIME] NULL,
 [DT_DATA]                      [SMALLDATETIME] NULL,
 [DT_EMISSAO]                   [SMALLDATETIME] NULL,
 [DT_GRAVACAO]                  [SMALLDATETIME] NULL,
 [DT_NFSUBSTITUTA]              [SMALLDATETIME] NULL,
 [DT_PERIFATFIM]                [SMALLDATETIME] NULL,
 [DT_PERIFATINI]                [SMALLDATETIME] NULL,
 [DT_VENCIMENTO]                [VARCHAR](20) NULL,
 [END_CLIENTE]                  [VARCHAR](100) NULL,
 [END_FILIAL]                   [VARCHAR](80) NULL,
 [FAX_FILIAL]                   [VARCHAR](18) NULL,
 [FLG_AVULSA]                   [BIT] NULL,
 [FLG_CANCELAMENTO]             [CHAR](1) NULL,
 [FLG_CONTROLEIMPRESSA]         [CHAR](1) NULL,
 [FLG_DADOSSUBCLIENTENF]        [CHAR](1) NULL,
 [FLG_DADOSSUBCLIENTESACADO]    [CHAR](1) NULL,
 [FLG_ELETRONICO]               [CHAR](1) NULL,
 [FLG_ENVIADA]                  [CHAR](1) NULL,
 [FLG_IMPRESSA]                 [BIT] NULL,
 [FLG_NFCOMPLEMENTAR]           [CHAR](1) NULL,
 [IC_CONTRIBUINTEESTADUAL]      [VARCHAR](1) NULL,
 [IC_SUBSTITUICAOTRIBUTARIA]    [CHAR](1) NULL,
 [IC_SUBSTITUICAOTRIBUTARIAISS] [CHAR](1) NULL,
 [INSCE_CLIENTE]                [VARCHAR](20) NULL,
 [INSCE_FILIAL]                 [VARCHAR](25) NULL,
 [INSCM_CLIENTE]                [VARCHAR](20) NULL,
 [INSCM_FILIAL]                 [VARCHAR](25) NULL,
 [LOTEFORMULARIO]               [VARCHAR](250) NULL,
 [MSG_FISCAL]                   [VARCHAR](200) NULL,
 [NOM_ABVD_FILIAL]              [CHAR](7) NULL,
 [NOM_CIDADE]                   [VARCHAR](100) NULL,
 [NOM_CLIENTE]                  [VARCHAR](100) NULL,
 [NOM_UF]                       [CHAR](2) NULL,
 [NOM_UFCOBRANCA]               [CHAR](2) NULL,
 [NOMEDAGRAFICA]                [VARCHAR](250) NULL,
 [NR_CONEXAO]                   [FLOAT] NULL,
 [NUM_FORMULARIO1]              [VARCHAR](15) NULL,
 [NUM_FORMULARIO2]              [VARCHAR](15) NULL,
 [NUM_NF]                       [VARCHAR](20) NULL,
 [PC_RETENCAIOINSS]             [NUMERIC](18, 0) NULL,
 [REGIMEESPECIALESTADO]         [VARCHAR](200) NULL,
 [REGIMEESPECIALMUNICIPIO]      [VARCHAR](200) NULL,
 [SEQ_CLIENTE]                  [CHAR](2) NULL,
 [SEQ_SUBCLIENTENF]             [CHAR](2) NULL,
 [SEQ_SUBCLIENTESACADO]         [CHAR](2) NULL,
 [SEQ_UTILIZA]                  [INT] NULL,
 [TEL_FILIAL]                   [VARCHAR](18) NULL,
 [TP_TRIBUTACAO]                [INT] NULL,
 [TOT_NOTA]                     [FLOAT] NULL,
 [TP_SERVICO]                   [CHAR](3) NULL,
 [TRB_ICMS]                     [VARCHAR](50) NULL,
 [TRB_ISS]                      [VARCHAR](25) NULL,
 [TX_ICMS]                      [FLOAT] NULL,
 [TX_ISS]                       [FLOAT] NULL,
 [UF_FILIAL]                    [CHAR](2) NULL,
 [UF_NOTA]                      [CHAR](2) NULL,
 [VAL_ICMS]                     [FLOAT] NULL,
 [VAL_INSS]                     [FLOAT] NULL,
 [VAL_INSS1]                    [VARCHAR](200) NULL,
 [VAL_INSS2]                    [VARCHAR](200) NULL,
 [VAL_INSS4]                    [VARCHAR](200) NULL,
 [VAL_INSS5]                    [VARCHAR](200) NULL,
 [VAL_ISS]                      [FLOAT] NULL,
 [DES_PRODUTO_ISS]              [VARCHAR](100) NULL,
 [DES_PRODUTO_ICMS]             [VARCHAR](100) NULL,
 [COD_ISS]                      [CHAR](5) NULL,
 [TX_ISS_TRIB]                  [FLOAT] NULL,
 [TX_ICMS_TRIB]                 [FLOAT] NULL,
 [ALIQ_PIS]                     [DECIMAL](18, 2) NULL,
 [ALIQ_COFINS]                  [DECIMAL](18, 2) NULL,
 [ALIQ_CSLL]                    [DECIMAL](18, 2) NULL,
 [ALIQ_IR]                      [DECIMAL](18, 2) NULL,
 [DT_VENCIMENTOAIDF]            [SMALLDATETIME] NULL,
 [DES_MSGALIQUOTA]              [VARCHAR](250) NULL,
 [DES_MSGELETRONICA]            [VARCHAR](200) NULL,
 [Flg_IsencaoISS]               [BIT] NULL,
 [Flg_IsencaoICMS]              [BIT] NULL,
 [Flg_IsencaoPIS]               [BIT] NULL,
 [Flg_IsencaoCOFINS]            [BIT] NULL,
 [Flg_IsencaoCSLL]              [BIT] NULL,
 [Flg_IsencaoIRRF]              [BIT] NULL,
 [Flg_Cepom]                    [BIT] NULL,
 [Des_MsgCepom]                 [VARCHAR](100) NULL,  
   
 --  <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>   
 [UF_ORIGEM]                    [CHAR](2) NULL,  
--  </alteracao>   
 [VAL_PIS]                      [FLOAT] NULL,
 [VAL_COFINS]                   [FLOAT] NULL,
 [VAL_CSLL]                     [FLOAT] NULL,
 [id_parametroinstituicao]      [BIT] NULL,
 [uf_destino]                   [CHAR](2) NULL,
 [des_metodo]                   [VARCHAR](30) NULL,
 Des_CidOrigemCTe               VARCHAR(100) NULL,
 Cod_IBGEOrigemCTe              INT NULL,
 Des_CidDestinoCTe              VARCHAR(100) NULL,
 Cod_IBGEDestinoCTe             INT NULL,
 --WKF 255294 - Ricardo Gomes - Plus It
 Val_IRRF                       DECIMAL(18, 2) NULL
 --/WKF 255294 - Ricardo Gomes - Plus It
);
     CREATE TABLE #SETOR
(CGC_CLIENTE           VARCHAR(15),
 SEQ_CLIENTE           VARCHAR(2),
 COD_FILIAL            INT,
 COD_REGIONAL          INT,
 COD_CLIENTETRIBUTACAO INT,
 DES_SETOREXTERNO      VARCHAR(20)
);
     CREATE INDEX IDX_#SETOR ON #SETOR(CGC_CLIENTE, SEQ_CLIENTE, COD_CLIENTETRIBUTACAO);
     CREATE TABLE #DADOS
(COD_SEQ        INT IDENTITY,
 COD_REGIONAL   INT,
 COD_FILIAL     INT,
 COD_NF         INT,
 QTDE           FLOAT,
 UNIDADE        CHAR(4),
 DES_PRODUTO    CHAR(50),
 COD_TRIBUTACAO INT,
 VAL_BRUTO      FLOAT,
 COD_SERIENOTA  CHAR(2)
);
     CREATE TABLE #DADOSFORMULARIO
(COD_REGIONAL        INT,
 COD_FILIAL          INT,
 TIP_NOTAFISCAL      INT,
 DES_PADRAOAIDFEST   CHAR(10),
 NUM_AIDFESTADO      CHAR(15),
 DT_DEFERAIDFESTADO  SMALLDATETIME,
 DES_REGESPPADRAOEST CHAR(100),
 DES_REGESPCOMPLMEST CHAR(200),
 DES_PADRAOAIDFMUN   CHAR(10),
 NUM_AIDFMUN         CHAR(15),
 DT_DEFERAIDFMUN     SMALLDATETIME,
 DES_REGESPPADRAOMUN CHAR(100),
 DES_REGESPCOMPLMMUN CHAR(200),
 NUM_PAFS            CHAR(15),
 DT_DEFERPAFS        SMALLDATETIME,
 NOM_GRAFICA         CHAR(200),
 GRF_UF              CHAR(2),
 GRF_CNPJ            CHAR(20),
 GRF_IESTAD          CHAR(30),
 GRF_IMUNIC          CHAR(30),
 NUM_FORMULINI       INT,
 NUM_FORMULFIM       INT,
 SER_FORMUL          CHAR(5),
 AIDFUF              CHAR(2),
 AIDFMUNICIPIO       CHAR(50)
);
     CREATE TABLE #FAT_MOVIMENTO
(COD_FILIAL                     INT NOT NULL,
 COD_REGIONAL                   INT NOT NULL,
 DT_MOVIMENTO                   SMALLDATETIME NOT NULL,
 COD_MOVIMENTO                  CHAR(11) NULL,
 COD_SEQMOVIMENTO               INT NULL,
 COD_TPMOVIMENTO                INT NOT NULL,
 COD_PRODUTO                    SMALLINT NOT NULL,
 COD_NF                         INT NULL,
 CGC_CLIENTE                    CHAR(15) NOT NULL,
 SEQ_CLIENTE                    CHAR(2) NOT NULL,
 COD_CLIENTETRIBUTACAO          INT NULL,
 COD_PRECO                      INT NULL,
 CGC_ORIGEM                     CHAR(15) NOT NULL,
 SEQ_ORIGEM                     CHAR(2) NOT NULL,
 CGC_DESTINO                    CHAR(15) NULL,
 SEQ_DESTINO                    CHAR(2) NULL,
 HORA_INICIO                    CHAR(4) NULL,
 HORA_FIM                       CHAR(4) NULL,
 TEMPO_PERMANENCIA              INT NULL,
 VAL_ADPERMANENCIA              FLOAT NULL,
 VAL_ADVALOREM                  FLOAT NULL,
 QTD_PROCESSAMENTO              FLOAT NULL,
 COD_TRIBUTACAO                 INT NOT NULL,
 VAL_ISS                        FLOAT NULL,
 TX_ISS                         FLOAT NULL,
 FLG_SUBSTITUICAOISS            CHAR(1) NULL,
 VAL_ICMS2                      FLOAT NULL,
 VAL_ICMS3                      FLOAT NULL,
 VAL_ICMS5                      FLOAT NULL,
 VAL_ICMS2_NF                   FLOAT NULL,
 VAL_ICMS3_NF                   FLOAT NULL,
 VAL_ICMS5_NF                   FLOAT NULL,
 TX_ICMS                        FLOAT NULL,
 FLG_SUBSTITUICAOICMS           CHAR(1) NULL,
 VAL_COFINS                     FLOAT NULL,
 VAL_PIS                        FLOAT NULL,
 VAL_INSS                       FLOAT NULL,
 VAL_LIQUIDO                    FLOAT NOT NULL,
 VAL_BRUTO                      FLOAT NOT NULL,
 QTD_EMBARQUE                   INT NULL,
 COD_FILIALSERVICO              INT NOT NULL,
 COD_REGIONALSERVICO            INT NOT NULL,
 DES_MOVIMENTO                  CHAR(15) NULL,
 FLGTPTRANSP                    CHAR(1) NULL,
 COD_FUNC                       INT NOT NULL,
 DT_ENTRADA                     SMALLDATETIME NOT NULL,
 COD_ROTA                       CHAR(10) NULL,
 QT_DIASCUSTODIA                INT NULL,
 DT_SAIDA                       SMALLDATETIME NULL,
 COD_PRECOEVENTUAL              INT NULL,
 COD_TRIBUTACAOSICOM            SMALLINT NULL,
 VAL_PRECO                      FLOAT NULL,
 COD_REDUZIDO                   INT NULL,
 COD_GTVEMBARQUE                CHAR(11) NULL,
 DAT_FECHAMENTO                 CHAR(2) NULL,
 QTD_ATMATENDIDO                INT NULL,
 COD_GRUPOFRANQUIA              INT NULL,
 COD_FRANQUIA                   INT NULL,
 COD_FUNCALTERACAO              INT NULL,
 DT_ALTERACAO                   SMALLDATETIME NULL,
 COD_FUNCRECALCULO              INT NULL,
 DT_RECALCULO                   SMALLDATETIME NULL,
 FLG_HORASEXTRAS                CHAR(1) NULL,
 FLG_COMPLEMENTAR               CHAR(1) NULL,
 COD_SERIE                      CHAR(2) NULL,
 KM_PERCORRIDO                  INT NULL,
 COD_CONJUGADA                  INT NOT NULL,
 INDICE_CONJUGADA               FLOAT NULL,
 FLG_PRECODIGITADO              CHAR(1) NULL,
 COD_FUNCTRANSFERENCIA          INT NULL,
 DT_TRANSFERENCIA               SMALLDATETIME NULL,
 DES_TRANSFERENCIA              CHAR(15) NULL,
 DES_ALTERACAO                  VARCHAR(255) NULL,
 COD_MOTIVORETORNO              INT NULL,
 FLG_GTVDEFINITIVO              VARCHAR(1) NULL,
 ID_LOGFRANQUIA                 INT NULL,
 COD_SERIENOTA                  CHAR(2) NULL,
 ID_PROBLEMA                    INT NULL,
 ID_SOLUCAO                     INT NULL,
 COD_CHAMADOCLIENTE             VARCHAR(15) NULL,
 COD_MAQUINA                    INT NULL,
 COD_CLISERV                    VARCHAR(18) NULL,
 FLG_COMPLREAJ                  CHAR(1) NULL,
 ID_CHAMADO                     INT NULL,
 PC_REPASSEICMS                 DECIMAL(18, 0) NULL,
 IC_ISS                         CHAR(1) NULL,
 IC_ICMS                        CHAR(1) NULL,
 IC_CONFINS                     CHAR(1) NULL,
 IC_PIS                         CHAR(1) NULL,
 ID_LOGPROCESSOREFATURADO       INT NULL,
 ID_LOGPROCESSOREFATURADOORIGEM INT NULL,
 VAL_ADICNOTURNO                FLOAT NULL
);  
   
--FIM CREATE TABLES----------------------------------------------------------------------------------            
             
             
 -- RETORNA O ESTADO DA TRIBUTAÇÃO DA NOTA              
     SELECT T.COD_TRIBUTACAO,
            T.COD_CIDADE,
            CE.NOM_ABVD_ESTADO,
            CE.COD_ESTADO
     INTO #UF_NOTA
     FROM SIC_TRIBUTACAO T
          LEFT JOIN COR_CIDADE CC ON T.COD_CIDADE = CC.COD_CIDADE
          LEFT JOIN COR_ESTADO CE ON CC.COD_ESTADO = CE.COD_ESTADO;             
             
            
  --INSERÇÃO DE REGISTROS NA TABELA #DADOSFORMULARIO            
     INSERT INTO #DADOSFORMULARIO
            SELECT FNF.COD_REGIONAL,
                   FNF.COD_FILIAL,
                   FNF.TIP_NOTAFISCAL,
                   FNF.DES_PADRAOAIDFEST,
                   FNF.NUM_AIDFESTADO,
                   FNF.DT_DEFERAIDFESTADO,
                   FNF.DES_REGESPPADRAOEST,
                   FNF.DES_REGESPCOMPLMEST,
                   FNF.DES_PADRAOAIDFMUN,
                   FNF.NUM_AIDFMUNIC,
                   FNF.DT_DEFERAIDFMUNIC,
                   FNF.DES_REGESPPADRAOMUN,
                   FNF.DES_REGESPCOMPLMMUN,
                   FNF.NUM_PAFS,
                   FNF.DT_DEFERPAFS,
                   FGF.NOM_GRAFICA,
                   FGF.GRF_UF,
                   FGF.GRF_CNPJ,
                   FGF.GRF_IESTAD,
                   FGF.GRF_IMUNIC,
                   FNF.NUM_FORMULINI,
                   FNF.NUM_FORMULFIM,
                   FNF.SER_FORMULARIO,
                   FNF.AIDFUF,
                   FNF.AIDFMUNICIPIO
            FROM FAT_FORMULARIONF FNF
                 INNER JOIN FAT_GRAFICANF FGF ON FNF.COD_GRAFICA = FGF.COD_GRAFICA
                 INNER JOIN FAT_NOTA N ON N.COD_REGIONAL = FNF.COD_REGIONAL
                                          AND N.COD_FILIAL = FNF.COD_FILIAL
            WHERE FNF.COD_REGIONAL = @COD_REGIONAL
                  AND FNF.COD_FILIAL = @COD_FILIAL
                  AND N.COD_NF = @COD_NF
                  AND N.COD_SERIENOTA = @COD_SERIENOTA
                  AND FNF.FLG_UTILIZADO = 0           
          
--  FNF.NUM_FORMULINI <= N.NUM_FORMULARIO1 AND              
--  FNF.NUM_FORMULFIM >= N.NUM_FORMULARIO2             
            ORDER BY FNF.NUM_LOTE;
     SELECT TOP 1 @LOTEFORMULARIO =
(
    SELECT TOP 1 'LOTE FORMULÁRIO '+RTRIM(CONVERT(CHAR(11), NUM_FORMULINI))+' A '+RTRIM(CONVERT(CHAR(11), NUM_FORMULFIM))+' SÉRIE '+CASE
                                                                                                                                        WHEN SER_FORMUL IS NULL
                                                                                                                                        THEN 'CA'
                                                                                                                                        ELSE SER_FORMUL
                                                                                                                                    END+'- '+RTRIM(DES_PADRAOAIDFEST)+' '+RTRIM(NUM_AIDFESTADO)+'('+AIDFUF+') DEFERIDA EM: '+CONVERT(CHAR(10), DT_DEFERAIDFESTADO, 103)+'- '+RTRIM(DES_PADRAOAIDFMUN)+' '+RTRIM(NUM_AIDFMUN)+' (P.M.'+RTRIM(AIDFMUNICIPIO)+') DEFERIDA EM '+CONVERT(CHAR(10), DT_DEFERAIDFMUN, 103)+' PAFS: '+RTRIM(NUM_PAFS)+' DEFERIDA EM '+CONVERT(CHAR(10), DT_DEFERPAFS, 103)
    FROM #DADOSFORMULARIO
    WHERE TIP_NOTAFISCAL = 0
);              
             
 -- DEFINE OS TEXTOS QUE SERÃO IMPRESSOS NAS NFS LASER.              
 -- NOME DA GRÁFICA            
     SELECT TOP 1 @NOMEDAGRAFICA =
(
    SELECT TOP 1 RTRIM(NOM_GRAFICA)+' / '+GRF_UF+' - CNPJ '+RTRIM(GRF_CNPJ)+' - INS.EST : '+GRF_IESTAD
    FROM #DADOSFORMULARIO
);              
             
 -- REGIME ESPECIAL ESTADO              
     SELECT TOP 1 @REGIMEESPECIALESTADO =
(
    SELECT TOP 1 RTRIM(DES_REGESPPADRAOEST)+'-'+DES_REGESPCOMPLMEST
    FROM #DADOSFORMULARIO
);            
             
 -- REGIME ESPECIAL MUNICIPIO            
     SELECT TOP 1 @REGIMEESPECIALMUNICIPIO =
(
    SELECT TOP 1 RTRIM(DES_REGESPPADRAOMUN)+'-'+DES_REGESPCOMPLMMUN
    FROM #DADOSFORMULARIO
);
     INSERT INTO #SETOR
            SELECT N.CGC_CLIENTE,
                   N.SEQ_CLIENTE,
                   N.COD_FILIAL,
                   N.COD_REGIONAL,
                   N.COD_CLIENTETRIBUTACAO,
                   EX.DES_SETOREXTERNO
            FROM FAT_NOTA N
                 INNER JOIN SIC_CLIENTETRIBUTACAO CT ON CT.CGC_CLIENTE = N.CGC_CLIENTE
                                                        AND CT.SEQ_CLIENTE = N.SEQ_CLIENTE
                                                        AND CT.COD_CLIENTETRIBUTACAO = N.COD_CLIENTETRIBUTACAO
                 LEFT JOIN SIC_SETOREXTERNO EX ON CT.ID_SETOREXTERNO = EX.ID_SETOREXTERNO
            WHERE N.COD_REGIONAL = @COD_REGIONAL
                  AND N.COD_FILIAL = @COD_FILIAL
                  AND N.COD_NF = @COD_NF
                  AND N.COD_SERIENOTA = @COD_SERIENOTA;             
             
            
 --SE CLIENTE FOR OPTANTE DO SIMPLES, PEGA AS ALIQUOTAS DA TABELA "SIC_TRIBISSNACIONAL"            
     DECLARE @FLG_SIMPLES INT, @ALIQ_PIS FLOAT, @ALIQ_COFINS FLOAT, @ALIQ_CSLL FLOAT, @ALIQ_IR FLOAT;
     SELECT @FLG_SIMPLES = ISNULL(FLG_SIMPLESNACIONAL, 0),
            @ALIQ_PIS = ISNULL(T.ALIQ_PIS, 0),
            @ALIQ_COFINS = ISNULL(T.ALIQ_COFINS, 0),
            @ALIQ_CSLL = ISNULL(T.ALIQ_CSLL, 0),
            @ALIQ_IR = ISNULL(T.ALIQ_IR, 0)
     FROM FAT_NOTA N
          INNER JOIN SIC_TRIBISSNACIONAL T ON T.COD_FAMILIAISS = N.COD_FAMILIAISS
          INNER JOIN SIC_CLIENTE C ON C.CGC_CLIENTE = N.CGC_CLIENTE
                                      AND C.SEQ_CLIENTE = N.SEQ_CLIENTE
     WHERE N.COD_FILIAL = @COD_FILIAL
           AND N.COD_REGIONAL = @COD_REGIONAL
           AND N.COD_NF = @COD_NF
           AND N.COD_SERIENOTA = @COD_SERIENOTA
           AND C.FLG_SIMPLESNACIONAL = 1
           AND T.FLG_SITUACAO = 1;     --alteração Kaio 29/06     
            
            
 --SELECIONA O ESTADO DO SERVIÇO ORIGINAL            
     DECLARE @UF_ORIGSERVICO CHAR(3);
     SELECT TOP 1 @UF_ORIGSERVICO = ISNULL(SC.UF_SUBCLIENTE, '')
     FROM FAT_NOTA N
          INNER JOIN FAT_MOVIMENTO M ON M.COD_FILIAL = N.COD_FILIAL
                                        AND M.COD_REGIONAL = N.COD_REGIONAL
                                        AND M.COD_NF = N.COD_NF
                                        AND M.COD_SERIENOTA = N.COD_SERIENOTA
          INNER JOIN SIC_CLIENTETRIBUTACAO T ON N.CGC_CLIENTE = T.CGC_CLIENTE
                                                AND N.SEQ_CLIENTE = T.SEQ_CLIENTE
                                                AND N.COD_CLIENTETRIBUTACAO = T.COD_CLIENTETRIBUTACAO
          INNER JOIN SIC_SUBCLIENTE SC ON T.CGC_SUBCLIENTESACADO = SC.CGC_SUBCLIENTE
                                          AND T.SEQ_SUBCLIENTESACADO = SC.SEQ
     WHERE N.COD_FILIAL = @COD_FILIAL
           AND N.COD_REGIONAL = @COD_REGIONAL
           AND N.COD_NF = @COD_NF
           AND N.COD_SERIENOTA = @COD_SERIENOTA;             
            
--Cria temporária da tabela nota para agilizar consulta          
     SELECT N.*
     INTO #FAT_NOTA
     FROM FAT_NOTA N
     WHERE N.COD_FILIAL = @COD_FILIAL
           AND N.COD_REGIONAL = @COD_REGIONAL
           AND N.COD_NF = @COD_NF
           AND N.COD_SERIENOTA = @COD_SERIENOTA;             
              
              
 --Cria temporária da tabela movimento para agilizar consulta            
     INSERT INTO #FAT_MOVIMENTO
(COD_FILIAL,
 COD_REGIONAL,
 DT_MOVIMENTO,
 COD_MOVIMENTO,
 COD_SEQMOVIMENTO,
 COD_TPMOVIMENTO,
 COD_PRODUTO,
 COD_NF,
 CGC_CLIENTE,
 SEQ_CLIENTE,
 COD_CLIENTETRIBUTACAO,
 COD_PRECO,
 CGC_ORIGEM,
 SEQ_ORIGEM,
 CGC_DESTINO,
 SEQ_DESTINO,
 HORA_INICIO,
 HORA_FIM,
 TEMPO_PERMANENCIA,
 VAL_ADPERMANENCIA,
 VAL_ADVALOREM,
 QTD_PROCESSAMENTO,
 COD_TRIBUTACAO,
 VAL_ISS,
 TX_ISS,
 FLG_SUBSTITUICAOISS,
 VAL_ICMS2,
 VAL_ICMS3,
 VAL_ICMS5,
 VAL_ICMS2_NF,
 VAL_ICMS3_NF,
 VAL_ICMS5_NF,
 TX_ICMS,
 FLG_SUBSTITUICAOICMS,
 VAL_COFINS,
 VAL_PIS,
 VAL_INSS,
 VAL_LIQUIDO,
 VAL_BRUTO,
 QTD_EMBARQUE,
 COD_FILIALSERVICO,
 COD_REGIONALSERVICO,
 DES_MOVIMENTO,
 FLGTPTRANSP,
 COD_FUNC,
 DT_ENTRADA,
 COD_ROTA,
 QT_DIASCUSTODIA,
 DT_SAIDA,
 COD_PRECOEVENTUAL,
 COD_TRIBUTACAOSICOM,
 VAL_PRECO,
 COD_REDUZIDO,
 COD_GTVEMBARQUE,
 DAT_FECHAMENTO,
 QTD_ATMATENDIDO,
 COD_GRUPOFRANQUIA,
 COD_FRANQUIA,
 COD_FUNCALTERACAO,
 DT_ALTERACAO,
 COD_FUNCRECALCULO,
 DT_RECALCULO,
 FLG_HORASEXTRAS,
 FLG_COMPLEMENTAR,
 COD_SERIE,
 KM_PERCORRIDO,
 COD_CONJUGADA,
 INDICE_CONJUGADA,
 FLG_PRECODIGITADO,
 COD_FUNCTRANSFERENCIA,
 DT_TRANSFERENCIA,
 DES_TRANSFERENCIA,
 DES_ALTERACAO,
 COD_MOTIVORETORNO,
 FLG_GTVDEFINITIVO,
 ID_LOGFRANQUIA,
 COD_SERIENOTA,
 ID_PROBLEMA,
 ID_SOLUCAO,
 COD_CHAMADOCLIENTE,
 COD_MAQUINA,
 COD_CLISERV,
 FLG_COMPLREAJ,
 ID_CHAMADO,
 PC_REPASSEICMS,
 IC_ISS,
 IC_ICMS,
 IC_CONFINS,
 IC_PIS,
 ID_LOGPROCESSOREFATURADO,
 ID_LOGPROCESSOREFATURADOORIGEM,
 VAL_ADICNOTURNO
)
            SELECT M.COD_FILIAL,
                   M.COD_REGIONAL,
                   M.DT_MOVIMENTO,
                   M.COD_MOVIMENTO,
                   M.COD_SEQMOVIMENTO,
                   M.COD_TPMOVIMENTO,
                   M.COD_PRODUTO,
                   M.COD_NF,
                   M.CGC_CLIENTE,
                   M.SEQ_CLIENTE,
                   M.COD_CLIENTETRIBUTACAO,
                   M.COD_PRECO,
                   M.CGC_ORIGEM,
                   M.SEQ_ORIGEM,
                   M.CGC_DESTINO,
                   M.SEQ_DESTINO,
                   M.HORA_INICIO,
                   M.HORA_FIM,
                   M.TEMPO_PERMANENCIA,
                   M.VAL_ADPERMANENCIA,
                   M.VAL_ADVALOREM,
                   M.QTD_PROCESSAMENTO,
                   M.COD_TRIBUTACAO,
                   M.VAL_ISS,
                   M.TX_ISS,
                   M.FLG_SUBSTITUICAOISS,
                   M.VAL_ICMS2,
                   M.VAL_ICMS3,
                   M.VAL_ICMS5,
                   M.VAL_ICMS2_NF,
                   M.VAL_ICMS3_NF,
                   M.VAL_ICMS5_NF,
                   M.TX_ICMS,
                   M.FLG_SUBSTITUICAOICMS,
                   M.VAL_COFINS,
                   M.VAL_PIS,
                   M.VAL_INSS,
                   M.VAL_LIQUIDO,
                   M.VAL_BRUTO,
                   M.QTD_EMBARQUE,
                   M.COD_FILIALSERVICO,
                   M.COD_REGIONALSERVICO,
                   M.DES_MOVIMENTO,
                   M.FLGTPTRANSP,
                   M.COD_FUNC,
                   M.DT_ENTRADA,
                   M.COD_ROTA,
                   M.QT_DIASCUSTODIA,
                   M.DT_SAIDA,
                   M.COD_PRECOEVENTUAL,
                   M.COD_TRIBUTACAOSICOM,
                   M.VAL_PRECO,
                   M.COD_REDUZIDO,
                   M.COD_GTVEMBARQUE,
                   M.DAT_FECHAMENTO,
                   M.QTD_ATMATENDIDO,
                   M.COD_GRUPOFRANQUIA,
                   M.COD_FRANQUIA,
                   M.COD_FUNCALTERACAO,
                   M.DT_ALTERACAO,
                   M.COD_FUNCRECALCULO,
                   M.DT_RECALCULO,
                   M.FLG_HORASEXTRAS,
                   M.FLG_COMPLEMENTAR,
                   M.COD_SERIE,
                   M.KM_PERCORRIDO,
                   M.COD_CONJUGADA,
                   M.INDICE_CONJUGADA,
                   M.FLG_PRECODIGITADO,
                   M.COD_FUNCTRANSFERENCIA,
                   M.DT_TRANSFERENCIA,
                   M.DES_TRANSFERENCIA,
                   M.DES_ALTERACAO,
                   M.COD_MOTIVORETORNO,
                   M.FLG_GTVDEFINITIVO,
                   M.ID_LOGFRANQUIA,
                   M.COD_SERIENOTA,
                   M.ID_PROBLEMA,
                   M.ID_SOLUCAO,
                   M.COD_CHAMADOCLIENTE,
                   M.COD_MAQUINA,
                   M.COD_CLISERV,
                   M.FLG_COMPLREAJ,
                   M.ID_CHAMADO,
                   M.PC_REPASSEICMS,
                   M.IC_ISS,
                   M.IC_ICMS,
                   M.IC_CONFINS,
                   M.IC_PIS,
                   M.ID_LOGPROCESSOREFATURADO,
                   M.ID_LOGPROCESSOREFATURADOORIGEM,
                   M.VAL_ADICNOTURNO
            FROM FAT_MOVIMENTO M
            WHERE M.COD_FILIAL = @COD_FILIAL
                  AND M.COD_REGIONAL = @COD_REGIONAL
                  AND M.COD_NF = @COD_NF
                  AND M.COD_SERIENOTA = @COD_SERIENOTA;             
           
 --Cria temporária da tabela de quebra para agilizar consulta            
     SELECT T.*
     INTO #SIC_CLIENTETRIBUTACAO
     FROM #FAT_NOTA N
          INNER JOIN SIC_CLIENTETRIBUTACAO T ON N.CGC_CLIENTE = T.CGC_CLIENTE
                                                AND N.SEQ_CLIENTE = T.SEQ_CLIENTE
                                                AND N.COD_CLIENTETRIBUTACAO = T.COD_CLIENTETRIBUTACAO
     WHERE N.COD_FILIAL = @COD_FILIAL
           AND N.COD_REGIONAL = @COD_REGIONAL
           AND N.COD_NF = @COD_NF
           AND N.COD_SERIENOTA = @COD_SERIENOTA;            
            
--<WKF 255294 - Ricardo Gomes - Plus It>
--Busca as mensagens de IRRf e PCC
     SELECT @MSG_IRRF = TRIBEMP.DES_MSGIRRF,
            @MSG_PCC = TRIBEMP.DES_MSGALIQUOTA,
            @ALIQ_IRRFTRIMEMP = CASE WHEN @FLG_SIMPLES = 1 THEN CAST(@ALIQ_IR AS DECIMAL(5,2))
                                ELSE CAST(TRIBEMP.ALIQ_IR AS DECIMAL(5, 2)) END
     FROM #FAT_NOTA NOTA
          INNER JOIN COR_Filial FIL ON NOTA.COD_REGIONAL = FIL.Cod_Regional
                                       AND NOTA.Cod_Filial = FIL.Cod_Filial
          INNER JOIN SIC_CLIENTE CLI ON CLI.CGC_CLIENTE = NOTA.CGC_CLIENTE
                                        AND CLI.SEQ_CLIENTE = NOTA.SEQ_CLIENTE
          INNER JOIN SIC_TIPOINSTITUICAO INST ON CLI.ID_TIPOINSTITUICAO = INST.ID_TIPOINSTITUICAO
          LEFT JOIN SIC_GRUPOINSTITUICAO GRP ON INST.ID_GRUPOINSTITUICAO = GRP.ID_GRUPOINSTITUICAO
          INNER JOIN SIC_TRIBUTACAOEMPRESA TRIBEMP ON TRIBEMP.COD_FAMILIAISS = NOTA.Cod_FamiliaISS
                                                      AND FIL.Cod_Empresa = TRIBEMP.COD_EMPRESA
                                                      AND TRIBEMP.ID_GRUPOINSTITUICAO = GRP.ID_GRUPOINSTITUICAO;
--Busca as mensagens de INSS retido e não retido
      SELECT @MSG_INSS = des_MsgINSS, @MSG_INSSNRETIDO = DES_MSGINSSNRETIDO
       FROM
     SIC_TributacaoMsgINSS AS inss
     INNER JOIN #FAT_NOTA AS NOTA ON inss.COD_FAMILIAISS = NOTA.Cod_FamiliaISS
                                     AND inss.Cod_Tributacao = nota.Cod_Tributacao;
--Busca o valor mínimo do imposto INSS
      SELECT @VALORMINIMPOSTO= CAST(VAL_MINIMOIMPOSTO AS FLOAT) FROM FAT_CONFIGIMPOSTOSFEDERAIS
--</WKF 255294 - Ricardo Gomes - Plus It>
			
--INICIO FAT_NOTAIMPRESSAO-------------------------------------------------------------------------------------            

     INSERT INTO #FAT_NOTAIMPRESSAO
(COD_FILIAL,
 COD_REGIONAL,
 COD_NF,
 COD_SERIENOTA,
 CGC_CLIENTE,
 CGC_CLIENTE2,
 CGC_FILIAL,
 SEQ_CLIENTE,
 COD_CLIENTETRIBUTACAO,
 COD_TRIBUTACAO,
 TP_TRIBUTACAO,
 TRB_ICMS,
 BAS_ICMS,
 VAL_ICMS,
 TX_ICMS,
 TX_ICMS_TRIB,
 TRB_ISS,
 BAS_ISS,
 VAL_ISS,
 TX_ISS,
 TX_ISS_TRIB,
 VAL_INSS,
 VAL_INSS1,
 VAL_INSS2,
 VAL_INSS4,
 VAL_INSS5,
 DES_SUBSTRIBICMS,
 DES_SUBCLIENTE,
 CON_CORRENTESANTANDER,
 TOT_NOTA,
 NUM_NF,
 CEP_FILIAL,
 COD_CGC,
 MSG_FISCAL,
 COD_NATUREZA,
 DT_GRAVACAO,
 DT_EMISSAO,
 COD_FUNCEMISSAO,
 DT_VENCIMENTO,
 DES_IRRF,
 END_CLIENTE,
 DES_CIDADE,
 CIDADE_CLIENTE,
 CEP_CLIENTE,
 NOM_UF,
 FLG_CANCELAMENTO,
 DES_CANCELAMENTO,
 COD_FUNCCANCELAMENTO,
 DT_CANCELAMENTO,
 FLG_IMPRESSA,
 FLG_AVULSA,
 TP_SERVICO,
 DES_JUSTFORASEQ,
 DES_VALOREXTENSO,
 DES_NUMBOLETO,
 DES_MSGBOLETO,
 DES_MSGNF,
 COD_RECIBO,
 FLG_ENVIADA,
 FLG_ELETRONICO,
 FLG_NFCOMPLEMENTAR,
 COD_BARRABOLETOBANC,
 COD_LINHADIGBOLETOBANC,
 COD_FAMILIAISS,
 FLG_CONTROLEIMPRESSA,
 NUM_FORMULARIO1,
 NUM_FORMULARIO2,
 SEQ_UTILIZA,
 DT_PERIFATINI,
 DT_PERIFATFIM,
 COD_NFSUBSTITUTA,
 COD_FUNCNFSUBS,
 DT_NFSUBSTITUTA,
 DT_DATA,
 CEP_COBRANCA,
 COD_EMPRESA,
 COD_FISCAL,
 COD_FISCAL_ISS,
 COD_INTERESTADUAL,
 COD_INTERMUNICIPAL,
 COD_RAMO,
 DES_BAIRROCOBRANCA,
 DES_CIDADECOBRANCA,
 DES_ENDCOBRANCA,
 DES_MENSAGEMSUBS_ICMS,
 DES_MENSAGEMSUBS_ISS,
 DES_SETOREXTERNO,
 DES_TPSUBCLIENTE,
 END_FILIAL,
 TEL_FILIAL,
 FAX_FILIAL,
 INSCE_FILIAL,
 INSCM_FILIAL,
 FLG_DADOSSUBCLIENTENF,
 FLG_DADOSSUBCLIENTESACADO,
 IC_CONTRIBUINTEESTADUAL,
 IC_SUBSTITUICAOTRIBUTARIA,
 IC_SUBSTITUICAOTRIBUTARIAISS,
 LOTEFORMULARIO,
 CGC_SUBCLIENTESACADO,
 NOM_ABVD_FILIAL,
 NOM_CIDADE,
 NOM_CLIENTE,
 NOM_UFCOBRANCA,
 INSCE_CLIENTE,
 INSCM_CLIENTE,
 NOMEDAGRAFICA,
 PC_RETENCAIOINSS,
 REGIMEESPECIALESTADO,
 REGIMEESPECIALMUNICIPIO,
 SEQ_SUBCLIENTENF,
 SEQ_SUBCLIENTESACADO,
 UF_FILIAL,
 UF_NOTA,
 ALIQ_PIS,
 ALIQ_COFINS,
 ALIQ_CSLL,
 ALIQ_IR,
 DES_MSGALIQUOTA,
 DESTACAR_ISS,
 DESTACAR_ICMS,
 Flg_IsencaoISS,
 Flg_IsencaoICMS,
 Flg_IsencaoPIS,
 Flg_IsencaoCOFINS,
 Flg_IsencaoCSLL,
 Flg_IsencaoIRRF,
 Flg_Cepom,
 Des_MsgCepom
 ,      
  --  <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>    
 UF_ORIGEM
 ,  
  --  </alteracao>   
    
  --NOVA REGRA FEDERAL DE PCC   
 VAL_PIS,
 VAL_COFINS,
 VAL_CSLL,
 id_parametroinstituicao,
 uf_destino,
 des_metodo,
 Des_CidOrigemCTe,
 Cod_IBGEOrigemCTe,
 Des_CidDestinoCTe,
 Cod_IBGEDestinoCTe,
  --WKF 255942 - Ricardo Gomes - Plus It 
 Val_IRRF
  --/Wkf 255942 - Ricardo Gomes - Plus It
    
)
            SELECT DISTINCT
                   N.COD_FILIAL,
                   N.COD_REGIONAL,
                   N.COD_NF,
                   N.COD_SERIENOTA,            
                         
  --N.CGC_CLIENTE,            
                   (DBO.FN_FAT_FORMATACGC(C.COD_CNPJCPF)+'('+CAST(C.SEQ_CLIENTE AS CHAR(2))+') - '+CAST(N.COD_CLIENTETRIBUTACAO AS VARCHAR(10))) AS CGC_CLIENTE,            
                                    
  --CGC_CLIENTE2            
                   CASE
                       WHEN T.FLG_DADOSSUBCLIENTENF = 'N'
                       THEN DBO.FN_FAT_FORMATACGC(N.CGC_CLIENTE)
                       ELSE DBO.FN_FAT_FORMATACGC(T.CGC_SUBCLIENTENF)
                   END AS CGC_CLIENTE2,            
             
  --CGC_FILIAL            
  --CASE WHEN N.COD_REGIONAL = 6 THEN            
  -- CASE N.COD_FILIAL            
  -- WHEN 1 THEN            
  --  '60.860.087/0137-72' --CURITIBA            
  -- WHEN 14 THEN            
  --  '60.860.087/0133-49' --FOZ DO IGUAÇU            
  -- WHEN 13 THEN            
  --  '60.860.087/0131-87' --PATO BRANCO            
  -- WHEN 11 THEN            
  --  '60.860.087/0138-53' --PONTA GROSSA            
  -- WHEN 18 THEN            
  --  '60.860.087/0134-20' --WENCESLAU BRAZ            
  -- WHEN 3 THEN            
  --  '60.860.087/0130-04' --MARINGA            
  -- WHEN 4 THEN            
  --  '60.860.087/0129-62' --UMUARAMA            
  -- WHEN 21 THEN            
  --  '60.860.087/0135-00' --UNIAO DA VITORIA            
-- WHEN 16 THEN            
  --  '60.860.087/0141-59' --FLO            
  -- WHEN 20 THEN            
  --  '60.860.087/0140-78' --BLU            
  -- WHEN 17 THEN            
  --  '60.860.087/0139-34' --JOI            
  -- END            
  --ELSE            
                   DBO.FN_FAT_FORMATACGC(F.COD_CGC) AS CGC_FILIAL,          
  --END AS CGC_FILIAL,            
              
                   N.SEQ_CLIENTE,
                   N.COD_CLIENTETRIBUTACAO,
                   N.COD_TRIBUTACAO,
                   N.TIP_TRIBUTACAO,            
          
  --TRB_ICMS            
                   CASE
                       WHEN N.TIP_TRIBUTACAO = 2
                       THEN 'SERVIÇOS FORA DO MUNICÍPIO - ICMS'
                       ELSE CASE
                                WHEN N.TIP_TRIBUTACAO = 3
                                     OR N.TIP_TRIBUTACAO = 5
                                THEN 'SERVIÇOS FORA DO ESTADO - ICMS'
                            END
                   END AS TRB_ICMS,
                   N.BAS_ICMS,
                   N.VAL_ICMS,
                   N.TX_ICMS,            
              
  --N.TX_ICMS_TRIB,            
                   CASE
                       WHEN N.TIP_TRIBUTACAO = 2
                       THEN '2'
                       ELSE CASE
                                WHEN N.TIP_TRIBUTACAO = 3
                                THEN '3'
                                ELSE CASE
                                         WHEN N.TIP_TRIBUTACAO = 5
                                         THEN '5'
                                     END
                            END
                   END AS TX_ICMS_TRIB,            
                               
  --TRB_ISS            
                   CASE
                       WHEN N.TIP_TRIBUTACAO = 0
                       THEN 'SERVIÇOS INTERNACIONAIS'
                       ELSE CASE
                                WHEN N.BAS_ISS > 0
                                THEN 'SERVIÇOS NO MUNICÍPIO'
                            END
                   END AS TRB_ISS,
                   N.BAS_ISS,
                   N.VAL_ISS,
                   N.TX_ISS,            
  
  --N.TX_ISS_TRIB,            
                   CASE
                       WHEN N.TIP_TRIBUTACAO = 0
                       THEN '0'
                       ELSE CASE
                                WHEN N.BAS_ISS > 0
                                THEN '1'
                            END
                   END AS TX_ISS_TRIB,            
            
  --VAL_INSS            
                   CASE
                       WHEN T.PC_RETENCAIOINSS > 0
                       THEN TR.TAX_INSS
                       ELSE 0
                   END AS VAL_INSS,            
              
  --N.VAL_INSS1,            
                   CASE
                       WHEN N.VAL_INSS > 0
                       THEN CASE
                                --<WKF 255294 - Ricard Gomes - Plus It>
                                WHEN N.VAL_INSS >= @VALORMINIMPOSTO    
                                --WHEN N.VAL_INSS >= 10 
                                --<WKF 255294 - Ricard Gomes - Plus It>
                                THEN 
                                   --<WKF 255294 - Ricard Gomes - Plus It>
                                   REPLACE(@MSG_INSS,'<INSS%>', CAST(ISNULL(TR.TAX_INSS, 0) AS VARCHAR(10)))
                                  --'RETENÇÃO PARA SEGURIDADE SOCIAL - IN 971 DE 13/11/2009 - INSS - '+CAST(ISNULL(TR.TAX_INSS, 0) AS VARCHAR(10))+'% SOBRE '+CAST(CAST(ISNULL(N.BAS_ISS, 0) + ISNULL(N.BAS_ICMS, 0) AS DECIMAL(18, 2)) AS VARCHAR(10))+' NO VALOR DE '+CAST(CAST(ISNULL(N.VAL_INSS, 0) AS DECIMAL(18, 2)) AS VARCHAR(10))
                                   --</WKF 255294 - Ricard Gomes - Plus It>
                                ELSE
                                  --<WKF 255294 - Ricard Gomes - Plus It>
                                  REPLACE(@MSG_INSSNRETIDO,'<INSS%>', CAST(ISNULL(TR.TAX_INSS, 0) AS VARCHAR(10)))
                                  --'NÃO INCIDÊNCIA DOS '+CAST(ISNULL(TR.TAX_INSS, 0) AS VARCHAR(10))+'% DA SEGURIDADE SOCIAL - CONFORME IN 971 DE 13/11/2009'
                                  --/<WKF 255294 - Ricard Gomes - Plus It>
                            END
						--<WKF 255294 - Ricard Gomes - Plus It>
                      -- ELSE 'NÃO INCIDÊNCIA DA RETENÇÃO P/ SEGURIDADE SOCIAL - IN 971 DE 13/11/2009 - INSS'
					   ELSE REPLACE(@MSG_INSSNRETIDO,'<INSS%>', CAST(ISNULL(TR.TAX_INSS, 0) AS VARCHAR(10)))
					   --<WKF 255294 - Ricard Gomes - Plus It>
                   END AS VAL_INSS1,            
              
  --VAL_INSS2      --<WKF 255294 - Ricard Gomes - Plus It>   
                   NULL AS VAL_INSS2,
                   --CASE
                       --WHEN N.VAL_INSS > 0
                       --THEN CASE
                       --         WHEN N.VAL_INSS >= 10
                       --         THEN 'IN 971 DE 13/11/2009 - INSS - '+CAST(ISNULL(TR.TAX_INSS, 0) AS VARCHAR(10))+'% SOBRE '+CAST(T.PC_RETENCAIOINSS AS VARCHAR(10))+'%'
                       --         ELSE 'CONFORME IN 971 DE 13/11/2009'
                       --     END
                       --ELSE 'SOCIAL - IN 971 DE 13/11/2009 - INSS'
                       --</WKF 255294 - Ricard Gomes - Plus It>
                   --END AS VAL_INSS2,            
               
  --VAL_INSS4            
                   CASE
                       WHEN N.VAL_INSS > 0
                            AND N.VAL_INSS  >= @VALORMINIMPOSTO    
                       THEN N.VAL_INSS
                   END AS VAL_INSS4,            
              
  --VAL_INSS5            
                   CASE
                       WHEN N.VAL_INSS > 0
                            AND T.PC_RETENCAIOINSS > 0
                            AND N.VAL_INSS >= @VALORMINIMPOSTO    
                       THEN T.PC_RETENCAIOINSS
                   END AS VAL_INSS5,     
              
  --DES_SUBSTRIBICMS           
  -- SDM 11093 NPrado TalantFour Inicio         
/*  CASE WHEN ISNULL(N.TIP_TRIBUTACAO,0) <> 1 AND T.IC_SUBSTITUICAOTRIBUTARIA = 'S' AND ISNULL(N.VAL_ICMS,0) = 0 THEN            
   CASE WHEN ISNULL(N.BAS_ISS,0) > 0 AND T.IC_SUBSTITUICAOTRIBUTARIAISS = 'S' AND ISNULL(N.VAL_ISS,0) = 0 THEN            
    ISNULL(TR.DES_SUBSTRIBICMS,'') + CHAR(13) + ISNULL(TR.DES_SUBSTRIBISS,'')            
   ELSE            
    ISNULL(TR.DES_SUBSTRIBICMS,'')            
   END            
  ELSE            
   CASE WHEN ISNULL(N.BAS_ISS,0) > 0 AND T.IC_SUBSTITUICAOTRIBUTARIAISS = 'S' AND ISNULL(N.VAL_ISS,0) = 0 THEN            
    ISNULL(TR.DES_SUBSTRIBISS,'')            
   ELSE            
    ''            
   END            
  END AS DES_SUBSTRIBICMS,    */

--SDM 11673 NPrado tratar isenção de tributo conforme Matriz de Isenção Inicio        
                   CASE
                       WHEN dbo.FN_FAT_ISENCAOICMS(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) = 1
                       THEN ''
                       ELSE CASE
                                WHEN ISNULL(M.Flg_SubstituicaoIcms, 'N') = 'S'
                                THEN ST.Des_Mensagem
                                ELSE ''
                            END -- AS  DES_SUBSTRIBICMS,        
                   END AS DES_SUBSTRIBICMS,        
          
  --SDM 11673 NPrado Fim        
          
  -- SDM 11093 NPrado TalantFour FIM        
              
  --DES_SUBCLIENTE            
                   'REF: SERVIÇOS PRESTADOS PARA '+TV.DES_TPSUBCLIENTE AS DES_SUBCLIENTE,            
             
  --CON_CORRENTESANTANDER            
                   CASE
                       WHEN N.DT_EMISSAO >= '20090824'
                       THEN CASE LEFT(N.CGC_CLIENTE, 9)
                                WHEN '033066408'
                                THEN 'DADOS P/PAGTO.: BRINKS SEG.TRANSP.VALORES - B.REAL AG.0689-C/C:670441-8 CNPJ: 60.860.087/0001-07'
                                WHEN '090400888'
                                THEN 'DADOS P/PAGTO.: BRINKS SEG.TRANSP.VALORES - B.REAL AG.0689-C/C:670441-8 CNPJ: 60.860.087/0001-07'
                                WHEN '061472676'
                                THEN 'DADOS P/PAGTO.: BRINKS SEG.TRANSP.VALORES - B.REAL AG.0689-C/C:670441-8 CNPJ: 60.860.087/0001-07'
                            END
                   END AS CON_CORRENTESANTANDER,            
              
  --TOT_NOTA            
                   (CAST(N.BAS_ISS AS FLOAT) + CAST(N.BAS_ICMS AS FLOAT)) AS TOT_NOTA,            
              
  --NUM_NF            
                   RIGHT(CAST('00000000'+CAST(N.COD_NF AS VARCHAR) AS VARCHAR), 8) AS NUM_NF,
                   F.CEP_FILIAL,            
              
  --COD_CGC            
                   DBO.FN_FAT_FORMATACGC(F.COD_CGC) AS COD_CGC,            
              
  --MSG_FISCAL            
                   CASE N.TIP_TRIBUTACAO
                       WHEN 0
                       THEN 'PRESTAÇÃO DE SERVIÇOS - 7358'
                       WHEN 1
                       THEN 'PRESTAÇÃO DE SERVIÇOS - '+CAST(CASE PR.COD_FAMILIAISS
                                                                WHEN 1
                                                                THEN TR.COD_FISCALISSESCOLTA
                                                                WHEN 2
                                                                THEN TR.COD_FISCALISSPROCNUM
                                                                WHEN 3
                                                                THEN TR.COD_FISCALISSFLM
                                                                WHEN 4
                                                                THEN TR.COD_FISCALISSPROCCHEQUE
                                                                WHEN 5
                                                                THEN TR.COD_FISCALISSMAOOBRA
                                                                WHEN 6
                                                                THEN TR.COD_FISCALISSTRANSP
                                                                WHEN 7
                                                                THEN TR.COD_FISCALISSVIGILANCIA
                                                                WHEN 8
                                                                THEN TR.Cod_FiscalIssAgeCarga
                                                                WHEN 9
                                                                THEN TR.Cod_FiscalIssDsbAdu
                                                                WHEN 11
                                                                THEN TR.COD_FISCALPROCENVELOPE
                                                                WHEN 12
                                                                THEN TR.COD_FISCALISSLOCCOMPUSAFE -- SDM-11377: INCLUSÃO TRATAMENTO FAMÍLIA 12 - LOCAÇÃO COMPUSAFE  
                                                                WHEN 13
                                                                THEN TR.COD_FISCALISSCONSULTSEG -- SDM-12332: INCLUSÃO TRATAMENTO FAMÍLIA 13 - CONSULTORIA EM SEGURANÇA            
                                                            END AS VARCHAR(10))
                       WHEN 2
                       THEN CASE
                                WHEN N.BAS_ISS <> 0
                                THEN 'PRESTAÇÃO DE SERVIÇOS - '+CAST(CASE PR.COD_FAMILIAISS
                                                                         WHEN 1
                                                                         THEN TR.COD_FISCALISSESCOLTA
                                                                         WHEN 2
                                                                         THEN TR.COD_FISCALISSPROCNUM
                                                                         WHEN 3
                                                                         THEN TR.COD_FISCALISSFLM
                                                                         WHEN 4
                                                                         THEN TR.COD_FISCALISSPROCCHEQUE
                                                                         WHEN 5
                                                                         THEN TR.COD_FISCALISSMAOOBRA
                                                                         WHEN 6
                                                                         THEN TR.COD_FISCALISSTRANSP
                                                                         WHEN 7
                                                                         THEN TR.COD_FISCALISSVIGILANCIA
                                                                         WHEN 8
                                                                         THEN TR.Cod_FiscalIssAgeCarga
                                                                         WHEN 9
                                                                         THEN TR.Cod_FiscalIssDsbAdu
                                                                         WHEN 11
                                                                         THEN TR.COD_FISCALPROCENVELOPE
                                                                         WHEN 12
                                                                         THEN TR.COD_FISCALISSLOCCOMPUSAFE -- SDM-11377: INCLUSÃO TRATAMENTO FAMÍLIA 12 - LOCAÇÃO COMPUSAFE      
                                                                         WHEN 13
                                                                         THEN TR.COD_FISCALISSCONSULTSEG -- SDM-12332: INCLUSÃO TRATAMENTO FAMÍLIA 13 - CONSULTORIA EM SEGURANÇA             
                                                                     END AS VARCHAR(10))+' - '+CAST(GR.COD_INTERMUNICIPAL AS VARCHAR(10))
                                ELSE  --CAMPO EXISTENTE SOMENTE NO TABELA DE MOVIMENTO    
           --  <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>   
                                CASE ISNULL(N.COD_NATUREZA, '')  --dbo.FN_FAT_CalcCFOP  ( M.COD_REGIONAL, M.COD_FILIAL,M.Cod_Preco,0)   
                                    WHEN '6.932'
                                    THEN 'PRESTAÇÃO DE SERVIÇOS - 6932'
                                    WHEN '5.932'
                                    THEN 'PRESTAÇÃO DE SERVIÇOS - 5932'
                                    WHEN ''
                                    THEN  
           --  </alteracao>   
                                CASE(SELECT MAX(M.CGC_ORIGEM) CGC_ORIGEM
                                     FROM FAT_MOVIMENTO M
                                     WHERE N.COD_NF = M.COD_NF
                                           AND N.COD_FILIAL = M.COD_FILIAL
                                           AND N.COD_REGIONAL = M.COD_REGIONAL
                                           AND N.COD_SERIENOTA = M.COD_SERIENOTA
                                     GROUP BY M.COD_NF,
                                              M.COD_FILIAL,
                                              M.COD_REGIONAL,
                                              M.COD_SERIENOTA)
                                    WHEN '004222428000130'
                                    THEN CASE
                                             WHEN ISNULL(M.Flg_SubstituicaoIcms, 'N') = 'N'
                                             THEN 'PRESTAÇÃO DE SERVIÇOS - 5932'
                                         END
                                    WHEN '020346524000146'
                                    THEN 'PRESTAÇÃO DE SERVIÇOS - 5932'
                                    ELSE 'PRESTAÇÃO DE SERVIÇOS - '+CAST(GR.COD_INTERMUNICIPAL AS VARCHAR(10))
                                END   
        --  <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>   
                                    ELSE 'PRESTAÇÃO DE SERVIÇOS - '+CAST(N.COD_NATUREZA AS VARCHAR(10))
                                END   
        --  </alteracao>   
          
                            END
                       WHEN 3
                       THEN CASE
                                WHEN N.BAS_ISS <> 0
                                THEN 'PRESTAÇÃO DE SERVIÇOS - '+CAST(CASE PR.COD_FAMILIAISS
                                                                         WHEN 1
                                                                         THEN TR.COD_FISCALISSESCOLTA
                                                                         WHEN 2
                                                                         THEN TR.COD_FISCALISSPROCNUM
                                                                         WHEN 3
                                                                         THEN TR.COD_FISCALISSFLM
                                                                         WHEN 4
                                                                         THEN TR.COD_FISCALISSPROCCHEQUE
                                                                         WHEN 5
                                                                         THEN TR.COD_FISCALISSMAOOBRA
                                                                         WHEN 6
                                                                         THEN TR.COD_FISCALISSTRANSP
                                                                         WHEN 7
                                                                         THEN TR.COD_FISCALISSVIGILANCIA
                                                                         WHEN 8
                                                                         THEN TR.Cod_FiscalIssAgeCarga
                                                                         WHEN 9
                                                                         THEN TR.Cod_FiscalIssDsbAdu
                                                                         WHEN 11
                                                                         THEN TR.COD_FISCALPROCENVELOPE
                                                                         WHEN 12
                                                                         THEN TR.COD_FISCALISSLOCCOMPUSAFE -- SDM-11377: INCLUSÃO TRATAMENTO FAMÍLIA 12 - LOCAÇÃO COMPUSAFE         
                                                                         WHEN 13
                                                                         THEN TR.COD_FISCALISSCONSULTSEG -- SDM-12332: INCLUSÃO TRATAMENTO FAMÍLIA 13 - CONSULTORIA EM SEGURANÇA         
                                                                     END AS VARCHAR(10))+' - '+CAST(GR.COD_INTERESTADUAL AS VARCHAR(10))
                                ELSE  --CAMPO EXISTENTE SOMENTE NO TABELA DE MOVIMENTO QUE NESSE MOMENTO NAO ESTÁ ALIMENTADA            
           --  <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>   
                                CASE ISNULL(N.COD_NATUREZA, '') -- dbo.FN_FAT_CalcCFOP  ( M.COD_REGIONAL, M.COD_FILIAL,M.Cod_Preco,0)   
                                    WHEN '6.932'
                                    THEN 'PRESTAÇÃO DE SERVIÇOS - 6932'
                                    WHEN '5.932'
                                    THEN 'PRESTAÇÃO DE SERVIÇOS - 5932'
                                    WHEN ''
                                    THEN     
           --  </alteracao>   
                                CASE(SELECT MAX(M.CGC_ORIGEM) CGC_ORIGEM
                                     FROM FAT_MOVIMENTO M
                                     WHERE N.COD_NF = M.COD_NF
                                           AND N.COD_FILIAL = M.COD_FILIAL
                                           AND N.COD_REGIONAL = M.COD_REGIONAL
                                           AND N.COD_SERIENOTA = M.COD_SERIENOTA
                                     GROUP BY M.COD_NF,
                                              M.COD_FILIAL,
                                              M.COD_REGIONAL,
                                              M.COD_SERIENOTA)
                                    WHEN '004222428000130'
                                    THEN CASE
                                             WHEN ISNULL(M.Flg_SubstituicaoIcms, 'N') = 'N'
                                             THEN 'PRESTAÇÃO DE SERVIÇOS - 6932'
                                         END
                                    WHEN '020346524000146'
                                    THEN 'PRESTAÇÃO DE SERVIÇOS - 6932'
                                    ELSE 'PRESTAÇÃO DE SERVIÇOS - '+CAST(GR.COD_INTERESTADUAL AS VARCHAR(10))
                                END   
        --  <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>  
                                    ELSE 'PRESTAÇÃO DE SERVIÇOS - '+CAST(N.COD_NATUREZA AS VARCHAR(10))
                                END   
        --  </alteracao>   
                            END
                       WHEN 5
                       THEN CASE
                                WHEN N.BAS_ISS <> 0
                                THEN 'PRESTAÇÃO DE SERVIÇOS - '+CAST(CASE PR.COD_FAMILIAISS
                                                                         WHEN 1
                                                                         THEN TR.COD_FISCALISSESCOLTA
                                                                         WHEN 2
                                                                         THEN TR.COD_FISCALISSPROCNUM
                                                                         WHEN 3
                                                                         THEN TR.COD_FISCALISSFLM
                                                                         WHEN 4
                                                                         THEN TR.COD_FISCALISSPROCCHEQUE
                                                                         WHEN 5
                                                                         THEN TR.COD_FISCALISSMAOOBRA
                                                                         WHEN 6
                                                                         THEN TR.COD_FISCALISSTRANSP
                                                                         WHEN 7
                                                                         THEN TR.COD_FISCALISSVIGILANCIA
                                                                         WHEN 8
                                                                         THEN TR.Cod_FiscalIssAgeCarga
                                                                         WHEN 9
                                                                         THEN TR.Cod_FiscalIssDsbAdu
                                                                         WHEN 11
                                                                         THEN TR.COD_FISCALPROCENVELOPE
                                                                         WHEN 12
                                                                         THEN TR.COD_FISCALISSLOCCOMPUSAFE -- SDM-11377: INCLUSÃO TRATAMENTO FAMÍLIA 12 - LOCAÇÃO COMPUSAFE         
                                                                         WHEN 13
                                                                         THEN TR.COD_FISCALISSCONSULTSEG -- SDM-12332: INCLUSÃO TRATAMENTO FAMÍLIA 13 - CONSULTORIA EM SEGURANÇA          
                                                                     END AS VARCHAR(10))+' - '+CAST(GR.COD_INTERESTADUAL AS VARCHAR(10))
                                ELSE --CAMPO EXISTENTE SOMENTE NO TABELA DE MOVIMENTO QUE NESSE MOMENTO NAO ESTÁ ALIMENTADA            
           --  <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>   
                                CASE ISNULL(N.COD_NATUREZA, '') -- dbo.FN_FAT_CalcCFOP  ( M.COD_REGIONAL, M.COD_FILIAL,M.Cod_Preco,0)   
                                    WHEN '6.932'
                                    THEN 'PRESTAÇÃO DE SERVIÇOS - 6932'
                                    WHEN '5.932'
                                    THEN 'PRESTAÇÃO DE SERVIÇOS - 5932'
                                    WHEN ''
                                    THEN    
           --  </alteracao>   
                                CASE(SELECT MAX(M.CGC_ORIGEM) CGC_ORIGEM
                                     FROM FAT_MOVIMENTO M
                                     WHERE N.COD_NF = M.COD_NF
                                           AND N.COD_FILIAL = M.COD_FILIAL
                                           AND N.COD_REGIONAL = M.COD_REGIONAL
                                           AND N.COD_SERIENOTA = M.COD_SERIENOTA
                                     GROUP BY M.COD_NF,
                                              M.COD_FILIAL,
                                              M.COD_REGIONAL,
                                              M.COD_SERIENOTA)
                                    WHEN '004222428000130'
                                    THEN CASE
                                             WHEN ISNULL(M.Flg_SubstituicaoIcms, 'N') = 'N'
                                             THEN 'PRESTAÇÃO DE SERVIÇOS - 6932'
                                         END
                                    WHEN '020346524000146'
                                    THEN 'PRESTAÇÃO DE SERVIÇOS - 6932'
                                    ELSE 'PRESTAÇÃO DE SERVIÇOS - '+CAST(GR.COD_INTERESTADUAL AS VARCHAR(10))
                                END   
        --  <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>   
                                    ELSE 'PRESTAÇÃO DE SERVIÇOS - '+CAST(N.COD_NATUREZA AS VARCHAR(10))
                                END   
        --  </alteracao>   
                            END
                       ELSE ''
                   END AS MSG_FISCAL,            
                     
  --COD_NATUREZA            
                   CASE N.TIP_TRIBUTACAO
                       WHEN 0
                       THEN '7358'
                       WHEN 1
                       THEN CASE
                                WHEN CE.NOM_ABVD_ESTADO = EN.NOM_ABVD_ESTADO
                                THEN '5.949'
                                ELSE '6.949'
                            END
                       ELSE            
        -- <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>   
                       CASE
                           WHEN N.TIP_TRIBUTACAO = 2
                                OR N.TIP_TRIBUTACAO = 3
                                OR N.TIP_TRIBUTACAO = 4
                                OR N.TIP_TRIBUTACAO = 5
                           THEN ISNULL(N.COD_NATUREZA, '') -- dbo.FN_FAT_CalcCFOP  ( M.COD_REGIONAL, M.COD_FILIAL,M.Cod_Preco,0)   
        --THEN CASE WHEN @UF_ORIGSERVICO <> CE.NOM_ABVD_ESTADO   
        --     THEN CASE WHEN CE.NOM_ABVD_ESTADO = EN.NOM_ABVD_ESTADO   
        --          THEN '5.932'            
        --          ELSE '6.932'            
        --          END            
        --     ELSE CASE WHEN CE.NOM_ABVD_ESTADO = EN.NOM_ABVD_ESTADO   
        --          THEN '5.' + RIGHT(GR.COD_INTERMUNICIPAL,3)            
        --          ELSE '6.' + RIGHT(GR.COD_INTERMUNICIPAL,3)            
        --          END            
        --     END    
        --  </alteracao>   
                           ELSE ''
                       END
                   END AS COD_NATUREZA,
                   N.DT_GRAVACAO,
                   N.DT_EMISSAO,
                   N.COD_FUNCEMISSAO,            
              
  --N.DT_VENCIMENTO,            
                   CASE
                       WHEN ISNULL(N.DT_VENCIMENTO, '') = ''
                       THEN 'C/ APRES'
                       ELSE CONVERT(VARCHAR(10), CAST(N.DT_VENCIMENTO AS DATETIME), 103)
                   END AS DT_VENCIMENTO,            
          
  --SDM 11673 NPrado verficar isenção do tributo conforme Matriz de Isenção Inicio.        
              
  --DES_IRRF         
                   CASE
                       WHEN dbo.FN_FAT_ISENCAOIRRF(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) = 1
                       THEN ''
                       ELSE CASE
                                WHEN(CAST(N.BAS_ISS AS FLOAT) + CAST(N.BAS_ICMS AS FLOAT)) >= 10
                                --<WKF 255294 - Ricardo Gomes - Plus It>
                                THEN REPLACE( @MSG_IRRF,'<IRRF%>', LTRIM(RTRIM( @ALIQ_IRRFTRIMEMP))) +
                                --THEN 'SUJEITO A RETENÇÃO DE '+                               
                                --  LTRIM(RTRIM(CAST(CASE 
                                --                   WHEN @FLG_SIMPLES = 1 THEN @ALIQ_IR
                                --                                         ELSE CAST(TRIBEMP.ALIQ_IR AS DECIMAL(18, 2))
                                --                                         END AS CHAR)))+
                                --  '% IRRF,ARTIGO 55, LEI 7713 DE 22/12/1988   -   R$ '+
                                --</WKF 255294 - Ricardo Gomes - Plus It>

                                  CAST(CAST(ROUND((CAST(ISNULL(N.BAS_ISS, 0) AS FLOAT) +
                                  CAST(ISNULL(N.BAS_ICMS, 0) AS FLOAT)) * 

                                --<WKF 255294 - Ricardo Gomes - Plus It>
                                  (@ALIQ_IRRFTRIMEMP/100), 2) AS DECIMAL(18, 2)) AS VARCHAR(20))
                                  END
                                --</WKF 255294 - Ricardo Gomes - Plus It>

                               --<WKF 255294 - Ricardo Gomes - Plus It>
                                --OBTÊM ALÍQUOTA  
                               --((CASE WHEN @FLG_SIMPLES = 1
                               --   THEN @ALIQ_IR
                               --   ELSE CAST(TRIBEMP.ALIQ_IR AS DECIMAL(18, 2)) END) / 100), 2) AS DECIMAL(18, 2)) AS VARCHAR(20))
                               --   END --AS DES_IRRF,  
                               --/<WKF 255294 - Ricardo Gomes - Plus It>
                               END AS DES_IRRF,           
            
  --SDM 11673 NPrado Fim        
             
  --END_CLIENTE            
                   CASE
                       WHEN T.FLG_DADOSSUBCLIENTENF = 'N'
                       THEN C.END_CLIENTE
                       ELSE T.DES_ENDCOBRANCA
                   END AS END_CLIENTE,            
              
  --DES_CIDADE            
                   CASE
                       WHEN T.FLG_DADOSSUBCLIENTENF = 'N'
                       THEN C.CIDADE_CLIENTE
                       ELSE T.DES_CIDADECOBRANCA
                   END AS DES_CIDADE,
                   ISNULL(C.CIDADE_CLIENTE, '') AS CIDADE_CLIENTE ,            
              
  --CEP_CLIENTE            
                   CASE
                       WHEN T.FLG_DADOSSUBCLIENTENF = 'N'
                       THEN C.CEP_CLIENTE
                       ELSE T.CEP_COBRANCA
                   END AS CEP_CLIENTE,            
              
  --NOM_UF            
                   CASE
                       WHEN T.FLG_DADOSSUBCLIENTENF = 'N'
                       THEN C.NOM_UF
                       ELSE T.NOM_UFCOBRANCA
                   END AS NOM_UF,
                   N.FLG_CANCELAMENTO,
                   N.DES_CANCELAMENTO AS DES_CANCELAMENTO,
                   N.COD_FUNCCANCELAMENTO AS COD_FUNCCANCELAMENTO,
                   N.DT_CANCELAMENTO AS DT_CANCELAMENTO,
                   N.FLG_IMPRESSA,
                   N.FLG_AVULSA,
                   N.TP_SERVICO,
                   N.DES_JUSTFORASEQ,
                   N.DES_VALOREXTENSO,
                   '' AS DES_NUMBOLETO,
                   '' AS DES_NUMBOLETO,
                   CASE
                       WHEN m.Flg_SubstituicaoIss = 'N'
                       THEN  --SDM 11974 NPrado 16/06/2014 - inicio      
        
                         CASE
                             WHEN dbo.FN_FAT_CADASTROCEPOM(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) = 1
                             THEN LEFT(isnull(TR.Des_MsgCepom, '')+' / '+LTRIM(RTRIM(N.DES_MSGNF)), 300)    --alteração Kaio - 05/01/2017   
                             ELSE LTRIM(RTRIM(N.DES_MSGNF))    --alteração Kaio - 05/01/2017      
                         END
                       ELSE LTRIM(RTRIM(N.DES_MSGNF))    --alteração Kaio - 05/01/2017   
                   END AS DES_MSGNF, --SDM 11974 NPrado 16/06/2014 - fim            
                   '' AS COD_RECIBO,
                   '' AS FLG_ENVIADA,
                   '' AS FLG_ELETRONICO,
                   N.FLG_NFCOMPLEMENTAR,
                   '' AS COD_BARRABOLETOBANC,
                   '' AS COD_LINHADIGBOLETOBANC,
                   N.COD_FAMILIAISS,
                   '' AS FLG_CONTROLEIMPRESSA,
                   '' AS NUM_FORMULARIO1,
                   '' AS NUM_FORMULARIO2,
                   '' AS SEQ_UTILIZA,
                   '' AS DT_PERIFATINI,
                   N.DT_PERIFATFIM,
                   '' AS COD_NFSUBSTITUTA,
                   '' AS COD_FUNCNFSUBS,
                   '' AS DT_NFSUBSTITUTA,
                   GETDATE() AS DT_DATA,
                   T.CEP_COBRANCA,
                   F.COD_EMPRESA,
                   F.COD_FISCAL,            
            
  --COD_FISCAL_ISS,              
                   CASE PR.COD_FAMILIAISS
                       WHEN 1
                       THEN TR.COD_FISCALISSESCOLTA
                       WHEN 2
                       THEN TR.COD_FISCALISSPROCNUM
                       WHEN 3
                       THEN TR.COD_FISCALISSFLM
                       WHEN 4
                       THEN TR.COD_FISCALISSPROCCHEQUE
                       WHEN 5
                       THEN TR.COD_FISCALISSMAOOBRA
                       WHEN 6
                       THEN TR.COD_FISCALISSTRANSP
                       WHEN 7
                       THEN TR.COD_FISCALISSVIGILANCIA
                       WHEN 8
                       THEN TR.Cod_FiscalIssAgeCarga
                       WHEN 9
                       THEN TR.Cod_FiscalIssDsbAdu
                       WHEN 11
                       THEN TR.COD_FISCALPROCENVELOPE
                       WHEN 12
                       THEN TR.COD_FISCALISSLOCCOMPUSAFE -- SDM-11377: INCLUSÃO TRATAMENTO FAMÍLIA 12 - LOCAÇÃO COMPUSAFE       
                       WHEN 13
                       THEN TR.COD_FISCALISSCONSULTSEG -- SDM-12332: INCLUSÃO TRATAMENTO FAMÍLIA 13 - CONSULTORIA EM SEGURANÇA          
                   END AS COD_FISCAL_ISS,
                   GR.COD_INTERESTADUAL,
                   GR.COD_INTERMUNICIPAL,
                   C.COD_RAMO,
                   T.DES_BAIRROCOBRANCA,
                   T.DES_CIDADECOBRANCA,
                   T.DES_ENDCOBRANCA,            
              
  --DES_MENSAGEMSUBS_ICMS         
  --SDM 11673 NPrado tratar mensagem conforme matriz de isenção Inicio.        
                   CASE
                       WHEN dbo.FN_FAT_ISENCAOICMS(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) = 1
                       THEN ''
                       ELSE
(
    SELECT TOP 1 TR2.DES_SUBSTRIBICMS
    FROM SIC_TRIBUTACAO TR2
    WHERE TR2.COD_TRIBUTACAO = N.COD_TRIBUTACAO
          AND N.TIP_TRIBUTACAO >= 2
) --AS DES_MENSAGEMSUBS_ICMS,            
                   END AS DES_MENSAGEMSUBS_ICMS,            
             
  --DES_MENSAGEMSUBS_ISS         
                   CASE
                       WHEN dbo.FN_FAT_ISENCAOISS(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) = 1
                       THEN ''
                       ELSE CASE
                                WHEN ISNULL(M.FLG_SUBSTITUICAOISS, 'N') = 'S'
                                THEN
(
    SELECT TOP 1 TR3.DES_SUBSTRIBISS
    FROM SIC_TRIBUTACAO TR3
    WHERE TR3.COD_TRIBUTACAO = N.COD_TRIBUTACAO
          AND (N.TIP_TRIBUTACAO = 1
               OR (N.TIP_TRIBUTACAO >= 2
                   AND N.BAS_ISS > 0))
)
                                ELSE ''
                            END -- AS DES_MENSAGEMSUBS_ISS,            
                   END AS DES_MENSAGEMSUBS_ISS,     
          
 --SDM 11673 NPrado tratar mensagem conforme matriz de isenção Fim.        
                
                   S.DES_SETOREXTERNO,
                   TV.DES_TPSUBCLIENTE,
                   L.DES_LOG+' '+F.END_FILIAL,
                   F.TEL_FILIAL,
                   F.FAX_FILIAL,
                   F.COD_INSCE,
                   F.COD_INSCM,
                   T.FLG_DADOSSUBCLIENTENF,
                   T.FLG_DADOSSUBCLIENTESACADO,
                   T.IC_CONTRIBUINTEESTADUAL,            
  --T.IC_SUBSTITUICAOTRIBUTARIA, SDM 11093         
  --DBO.FN_FAT_SUBSTRIBUTARIAICMS ( M.Cod_Preco, M.COD_REGIONAL, M.COD_FILIAL), --SDM 11093       
                   ISNULL(M.Flg_SubstituicaoIcms, 'N'), --SDM 11673       
                   M.FLG_SUBSTITUICAOISS,            
            
  --FORMULÁRIOS (RODAPÉ DA NF)              
                   @LOTEFORMULARIO AS LOTEFORMULARIO,
                   T.CGC_SUBCLIENTESACADO,
                   F.NOM_ABVD_FILIAL,
                   CI.NOM_CIDADE,
                   C.NOM_CLIENTE,
                   T.NOM_UFCOBRANCA,
                   C.COD_INSCE,
                   C.COD_INSCM,
                   @NOMEDAGRAFICA AS NOMEDAGRAFICA,
                   T.PC_RETENCAIOINSS,
                   @REGIMEESPECIALESTADO AS REGIMEESPECIALESTADO,
                   @REGIMEESPECIALMUNICIPIO AS REGIMEESPECIALMUNICIPIO,
                   T.SEQ_SUBCLIENTENF,
                   T.SEQ_SUBCLIENTESACADO,
                   CE.NOM_ABVD_ESTADO UF_FILIAL,
                   EN.NOM_ABVD_ESTADO UF_NOTA,
                   CASE
                       WHEN @FLG_SIMPLES = 1
                       THEN @ALIQ_PIS
                       ELSE CAST(TRIBEMP.ALIQ_PIS AS DECIMAL(18, 2))
                   END AS ALIQ_PIS,
                   CASE
                       WHEN @FLG_SIMPLES = 1
                       THEN @ALIQ_COFINS
                       ELSE CAST(TRIBEMP.ALIQ_COFINS AS DECIMAL(18, 2))
                   END AS ALIQ_COFINS,
                   CASE
                       WHEN @FLG_SIMPLES = 1
                       THEN @ALIQ_CSLL
                       ELSE CAST(TRIBEMP.ALIQ_CSLL AS DECIMAL(18, 2))
                   END AS ALIQ_CSLL,
                   CASE
                       WHEN @FLG_SIMPLES = 1
                       THEN @ALIQ_IR
                       ELSE CAST(TRIBEMP.ALIQ_IR AS DECIMAL(18, 2))
                   END AS ALIQ_IR,            
              
  --DES_MSGALIQUOTA            
  --SDM 11673 NPrado verificar se existe isenção de tributo conf. Matriz de Isenção Inicio.        
-- CASE WHEN dbo.FN_FAT_ISENCAOPIS ( M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) = 1 THEN        
--   ''        
-- ELSE           
--    CASE WHEN LTRIM(RTRIM(ISNULL(GRUP.DES_GRUPOINSTITUICAO,''))) = 'ORGAO PUBLICO' THEN            
--     'PIS: ' + CAST(CAST((ISNULL(N.BAS_ISS,0) + ISNULL(N.BAS_ICMS,0)) * CASE WHEN @FLG_SIMPLES = 1 THEN (@ALIQ_PIS/100) ELSE (ISNULL(TRIBEMP.ALIQ_PIS,0)/100) END AS DECIMAL(18,2)) AS VARCHAR(10))        
--      + '  COFINS: ' + CAST(CAST((ISNULL(N.BAS_ISS,0) + ISNULL(N.BAS_ICMS,0)) * CASE WHEN @FLG_SIMPLES = 1 THEN (@ALIQ_COFINS/100) ELSE (ISNULL(TRIBEMP.ALIQ_COFINS,0)/100) END AS DECIMAL(18,2)) AS VARCHAR(10)) + '  CSLL: ' + CAST(CAST((ISNULL(N.BAS_ISS
--                   0,   
     
     
--)         
--  + ISNULL(N.BAS_ICMS,0)) * CASE WHEN @FLG_SIMPLES = 1 THEN (@ALIQ_CSLL/100) ELSE (ISNULL(TRIBEMP.ALIQ_CSLL,0)/100) END AS DECIMAL(18,2)) AS VARCHAR(10))            
--    ELSE            
--     CASE WHEN (CAST(ISNULL(N.BAS_ISS,0) AS FLOAT) + CAST(ISNULL(N.BAS_ICMS,0) AS FLOAT)) > 5000 THEN            
--      'PIS: ' + CAST(CAST((ISNULL(N.BAS_ISS,0) + ISNULL(N.BAS_ICMS,0)) * CASE WHEN @FLG_SIMPLES = 1 THEN (@ALIQ_PIS/100) ELSE (ISNULL(TRIBEMP.ALIQ_PIS,0)/100) END AS DECIMAL(18,2)) AS VARCHAR(10))        
--       + '  COFINS: ' + CAST(CAST((ISNULL(N.BAS_ISS,0) + ISNULL(N.BAS_ICMS,0)) * CASE WHEN @FLG_SIMPLES = 1 THEN (@ALIQ_COFINS/100) ELSE (ISNULL(TRIBEMP.ALIQ_COFINS,0)/100) END AS DECIMAL(18,2)) AS VARCHAR(10)) + '  CSLL: ' + CAST(CAST((ISNULL(N.BAS_ISS
--                   0,
    
      
--)         
--  + ISNULL(N.BAS_ICMS,0)) * CASE WHEN @FLG_SIMPLES = 1 THEN (@ALIQ_CSLL/100) ELSE (ISNULL(TRIBEMP.ALIQ_CSLL,0)/100) END AS DECIMAL(18,2)) AS VARCHAR(10))            
--     ELSE            
--      'SUJEITO A RETENÇÃO DE ' +             
--      CASE WHEN @FLG_SIMPLES = 1 THEN            
--       CAST(CAST((@ALIQ_PIS + @ALIQ_COFINS + @ALIQ_CSLL) AS DECIMAL(18,2)) AS VARCHAR(20))            
--      ELSE            
--       CAST((ISNULL(TRIBEMP.ALIQ_PIS,0) + ISNULL(TRIBEMP.ALIQ_COFINS,0) + ISNULL(TRIBEMP.ALIQ_CSLL,0)) AS VARCHAR(20))            
--      END +             
--      ' DE PIS/COFINS/CSLL CONF. LEI 10833 ART 30 E 31 DE 29/12/2003'            
--    END            
--    END --AS DES_MSGALIQUOTA,          
--   END AS DES_MSGALIQUOTA,          
  
  --Será calculado em separado e não no momento da Query, procurando facilitar no código e possivelmente o desempenho  
                   '' AS DES_MSGALIQUOTA,   
  
  
   --SDM 11673 NPrado Fim        
           
 --DESTACAR_ISS          
                   T.IC_DESTACARISS,          
 --DESTACAR_ICMS          
                   T.IC_DESTACARICMS,         
 --SDM 11673 NPrado - Inicio            
                   dbo.FN_FAT_ISENCAOISS(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) AS Flg_IsencaoISS,
                   dbo.FN_FAT_ISENCAOICMS(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) AS Flg_IsencaoICMS,
                   dbo.FN_FAT_ISENCAOPIS(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) AS Flg_IsencaoPIS,
                   dbo.FN_FAT_ISENCAOCOFINS(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) AS Flg_IsencaoCOFINS,
                   dbo.FN_FAT_ISENCAOCSLL(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) AS Flg_IsencaoCSLL,
                   dbo.FN_FAT_ISENCAOIRRF(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) AS Flg_IsencaoIRRF,        
   
   --SDM 11974 NPrado - Inicio           
                   CASE
                       WHEN m.Flg_SubstituicaoIss = 'N'
                       THEN dbo.FN_FAT_CADASTROCEPOM(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO)
                       ELSE 0
                   END AS Flg_Cepom,
                   CASE
                       WHEN m.Flg_SubstituicaoIss = 'N'
                       THEN CASE
                                WHEN dbo.FN_FAT_CADASTROCEPOM(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) = 1
                                THEN TR.Des_MsgCepom
                                ELSE ''
                            END
                   END AS Des_MsgCepom,      
   
 --SDM 11673 Nprado - Fim        
--  <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>    
                   N.UF_ORIGEM,  
--  </alteracao>   
  
  --NOVA REGRA FEDERAL DE PCC   
                   0, --VAL_PIS   
                   0, --VAL_COFINS   
                   0, --VAL_CSLL   
                   N.ID_PARAMETROINSTITUICAO,
                   N.UF_DESTINO,
                   (CASE
                        WHEN T.TP_COBRANCA = 'B'
                        THEN IC.DES_METODO
                        ELSE NULL
                    END) AS DES_METODO,
                   N.Des_CidOrigemCTe,
                   N.Cod_IBGEOrigemCTe,
                   N.Des_CidDestinoCTe,
                   N.Cod_IBGEDestinoCTe,  
  --WKF 255294 - Ricardo Gomes - Plus It
  --Calcula o valor do IRRF 
                   CAST(CAST(ROUND((CAST(ISNULL(N.BAS_ISS, 0) AS FLOAT) + CAST(ISNULL(N.BAS_ICMS, 0) AS FLOAT)) * --0.01  
  
            --OBTÊM ALÍQUOTA  
                                   ((CASE
                                         WHEN @FLG_SIMPLES = 1
                                         THEN @ALIQ_IR
                                         ELSE CAST(TRIBEMP.ALIQ_IR AS DECIMAL(18, 2))
                                     END) / 100), 2) AS DECIMAL(18, 2)) AS DECIMAL(18, 2)) AS VAL_IRRF
  --/WKF 255942 - Ricardo Gomes - PLus It
            FROM #FAT_NOTA N
                 INNER JOIN #SIC_CLIENTETRIBUTACAO T ON N.CGC_CLIENTE = T.CGC_CLIENTE
                                                        AND N.SEQ_CLIENTE = T.SEQ_CLIENTE
                                                        AND N.COD_CLIENTETRIBUTACAO = T.COD_CLIENTETRIBUTACAO
                 INNER JOIN SIC_TRIBUTACAO TR ON TR.COD_TRIBUTACAO = N.COD_TRIBUTACAO
                 LEFT JOIN #UF_NOTA EN ON TR.COD_CIDADE = EN.COD_CIDADE
                                          AND EN.COD_TRIBUTACAO = TR.COD_TRIBUTACAO
                 LEFT JOIN COR_TIPOSUBCLIENTE TV ON T.COD_TPSUBCLIENTE = TV.COD_TPSUBCLIENTE
                 LEFT JOIN COR_FILIAL F ON N.COD_REGIONAL = F.COD_REGIONAL
                                            AND N.COD_FILIAL = F.COD_FILIAL              
--SDM 13333 - INICIO  
                 INNER JOIN COR_EMPRESA E(NOLOCK) ON F.COD_EMPRESA = E.COD_EMPRESA
                 INNER JOIN FAT_PARAMETROINSTITUICAO INS(NOLOCK) ON E.COD_EMPRESA = INS.COD_EMPRESA
                 INNER JOIN FAT_INSTITUICAOCOBRANCA IC(NOLOCK) ON INS.ID_INSTITUICAOCOBRANCA = IC.ID_INSTITUICAOCOBRANCA   
--SDM 13333 - FIM  
                 INNER JOIN COR_LOGRADOURO L ON L.COD_LOG = F.COD_LOG
                 INNER JOIN COR_CIDADE CI ON F.COD_CIDADE = CI.COD_CIDADE
                 INNER JOIN COR_ESTADO CE ON CI.COD_ESTADO = CE.COD_ESTADO
                 INNER JOIN SIC_CLIENTE C ON C.CGC_CLIENTE = N.CGC_CLIENTE
                                             AND C.SEQ_CLIENTE = N.SEQ_CLIENTE               
 --INNER JOIN          
 -- SIC_SUBCLIENTE SC (NOLOCK)              
 --ON              
 -- T.CGC_SUBCLIENTESACADO = SC.CGC_SUBCLIENTE AND              
 -- T.SEQ_SUBCLIENTESACADO = SC.SEQ            
                 INNER JOIN SIC_RAMO_ATIVIDADE RA ON C.COD_RAMO = RA.COD_RAMO
                 LEFT JOIN #FAT_MOVIMENTO M ON N.COD_NF = M.COD_NF
                                               AND N.COD_FILIAL = M.COD_FILIAL
                                               AND N.COD_REGIONAL = M.COD_REGIONAL
                                               AND N.COD_SERIENOTA = M.COD_SERIENOTA
                 INNER JOIN SIC_GRUPORAMOATIVIDADE GR --ON RA.COD_GRUPORAMO = GR.COD_GRUPORAMO           
                 ON GR.COD_GRUPORAMO = CASE
                                           WHEN ISNULL(M.Flg_SubstituicaoIcms, 'N') = 'N'
                                           THEN RA.COD_GRUPORAMO
                                           ELSE 10
                                       END
                 LEFT JOIN #SETOR S ON S.CGC_CLIENTE = N.CGC_CLIENTE
                                       AND S.SEQ_CLIENTE = N.SEQ_CLIENTE
                                       AND S.COD_CLIENTETRIBUTACAO = N.COD_CLIENTETRIBUTACAO
                 LEFT JOIN SIC_PRODUTO PR ON PR.COD_PRODUTO = M.COD_PRODUTO
                 LEFT JOIN SIC_TIPOINSTITUICAO INST ON C.ID_TIPOINSTITUICAO = INST.ID_TIPOINSTITUICAO
                 LEFT JOIN SIC_GRUPOINSTITUICAO GRUP ON GRUP.ID_GRUPOINSTITUICAO = INST.ID_GRUPOINSTITUICAO
                 LEFT JOIN SIC_TRIBUTACAOEMPRESA TRIBEMP ON TRIBEMP.ID_GRUPOINSTITUICAO = GRUP.ID_GRUPOINSTITUICAO
                                                            AND TRIBEMP.COD_FAMILIAISS = PR.COD_FAMILIAISS
                                                            AND TRIBEMP.COD_EMPRESA = F.COD_EMPRESA
                 LEFT JOIN -- SDM 11093 NPrado TalentFour Inicio        
                 SIC_SUBSTTRIBUTICMS ST ON ST.UF_Estado = CE.Cod_Estado    -- SDM 11093 NPrado TalentFour Fim                 
            WHERE N.COD_FILIAL = @COD_FILIAL
                  AND N.COD_REGIONAL = @COD_REGIONAL
                  AND N.COD_NF = @COD_NF
                  AND N.COD_SERIENOTA = @COD_SERIENOTA;            
             
             
--FIM FAT_NOTAIMPRESSAO---------------------------------------------------------------------------------------            
       
 --********************** Cálculo de PIS\COFINS\CSLL  
   
   
  
  
 --********* 1º Verifica critérios para descobrir se deverá calcular PIS\COFINS\CSLL  

     SELECT DISTINCT
            dbo.FN_FAT_ISENCAOPIS(M.Cod_Preco, M.COD_REGIONALSERVICO, M.COD_FILIALSERVICO) AS ISENCAO
     INTO #ISENCAO_PIS
     FROM #FAT_NOTAIMPRESSAO N
          LEFT JOIN #FAT_MOVIMENTO M ON N.COD_NF = M.COD_NF
                                        AND N.COD_FILIAL = M.COD_FILIAL
                                        AND N.COD_REGIONAL = M.COD_REGIONAL
                                        AND N.COD_SERIENOTA = M.COD_SERIENOTA;      
   
   
   --OBTÊM ISENÇÃO  
     SELECT @BO_ISENTO_IMPOST_FEDERAIS = ISENCAO
     FROM #ISENCAO_PIS;  
  
   --FORMATA VALOR  
     IF @BO_ISENTO_IMPOST_FEDERAIS IS NULL
         BEGIN
             SET @BO_ISENTO_IMPOST_FEDERAIS = 0;
         END;  
  
  --ARMAZENA INFORMAÇÃO  
     SELECT @BO_CALCULA_IMPOST_FEDERAIS = CASE
                                              WHEN @BO_ISENTO_IMPOST_FEDERAIS = 1
                                              THEN 0
                                              ELSE 1
                                          END;  
    
     
   --SE NÃO POSSUI ISENÇÃO PROSSEGUE COM OS CÁLCULOS  
     IF @BO_CALCULA_IMPOST_FEDERAIS = 1
         BEGIN  
   --**** REALIZA CÁLCULO  
             UPDATE F
               SET
                   F.VAL_PIS = CAST(TOT_NOTA * (ALIQ_PIS / 100) AS DECIMAL(18, 2)),
                   F.VAL_COFINS = CAST(TOT_NOTA * (ALIQ_COFINS / 100) AS DECIMAL(18, 2)),
                   F.VAL_CSLL = CAST(TOT_NOTA * (ALIQ_CSLL / 100) AS DECIMAL(18, 2))
             FROM #FAT_NOTAIMPRESSAO F;  
  
    --**** VERIFICA SE A NOTA ATENDE OS CRITÉRIOS PARA CALCULAR  

             IF @FED_VAL_MINIMOIMPOSTO > 0
                 BEGIN  
      
     -- CRITÉRIO DE VALOR MÍNIMO DE IMPOSTO  
                     IF
(
    SELECT(VAL_PIS + VAL_COFINS + VAL_CSLL)
    FROM #FAT_NOTAIMPRESSAO
) <= @FED_VAL_MINIMOIMPOSTO
                         BEGIN  
       
      --ZERA IMPOSTOS  
                             SET @BO_CALCULA_IMPOST_FEDERAIS = 0;
                         END;
                 END;
                 ELSE --IF @FED_VAL_MINIMOIMPOSTO >0  
                 BEGIN  
     
     -- CRITÉRIO DE VALOR MÍNIMO DE NOTA  
                     IF
(
    SELECT TOT_NOTA
    FROM #FAT_NOTAIMPRESSAO
) <= @FED_VAL_MINIMONOTA
                         BEGIN  
       
      --ZERA IMPOSTOS  
                             SET @BO_CALCULA_IMPOST_FEDERAIS = 0;
                         END;
                 END; --IF @FED_VAL_MINIMOIMPOSTO >0  
     
   --Se está isento de calcular impostos federais verifica o critério de instituições que sempre cobram imposto  
   --IF @BO_CALCULA_IMPOST_FEDERAIS=0  
   --BEGIN  
     
   -- --VERIFICA SE A INSTITUIÇÃO COBRA O PIS\COFINS\CSLL  
   -- SELECT   
   --  @BO_CALCULA_IMPOST_FEDERAIS =  
   --   CASE WHEN  
   --     LTRIM(RTRIM(ISNULL(GRUP.DES_GRUPOINSTITUICAO,''))) = 'ORGAO PUBLICO'  
   --    THEN  
   --     1  
   --    ELSE   
   --     0  
   --   END  
         
  
   --  FROM             
   --   #FAT_NOTA N    
   --  INNER JOIN            
   --   SIC_CLIENTE C           
   --  ON C.CGC_CLIENTE = N.CGC_CLIENTE AND              
   --   C.SEQ_CLIENTE = N.SEQ_CLIENTE   
   --  LEFT JOIN             
   --   SIC_TIPOINSTITUICAO INST           
   --  ON      C.ID_TIPOINSTITUICAO = INST.ID_TIPOINSTITUICAO            
   --  LEFT JOIN             
   --   SIC_GRUPOINSTITUICAO GRUP           
   --  ON      GRUP.ID_GRUPOINSTITUICAO = INST.ID_GRUPOINSTITUICAO     
  
   --END   


         END; --IF @BO_CALCULA_IMPOST_FEDERAIS  
     
 --********* 2º Verifica critérios para descobrir se deverá calcular PIS\COFINS\CSLL  


     IF @BO_CALCULA_IMPOST_FEDERAIS = 1
         BEGIN  
       
   --**** REALIZA CÁLCULO  
             UPDATE F
               SET
                   F.VAL_PIS = CAST(TOT_NOTA * (ALIQ_PIS / 100) AS DECIMAL(18, 2)),
                   F.VAL_COFINS = CAST(TOT_NOTA * (ALIQ_COFINS / 100) AS DECIMAL(18, 2)),
                   F.VAL_CSLL = CAST(TOT_NOTA * (ALIQ_CSLL / 100) AS DECIMAL(18, 2))
             FROM #FAT_NOTAIMPRESSAO F;  
     
     
   --**** ACRESCENTA MENSAGEM  
            --<WKF 255294 - Ricardo Gomes - Plus It>

            SELECT @VAL_PCC_MSG = CAST( F.VAL_PIS  + F.VAL_COFINS + F.VAL_CSLL AS VARCHAR (20)),
              @IMPOSTO_PCC_MSG =  CAST(F.ALIQ_PIS + F.ALIQ_COFINS + F.ALIQ_CSLL AS VARCHAR(10))
              FROM #FAT_NOTAIMPRESSAO F;

            SET @MSG_PCC = REPLACE(@MSG_PCC + @VAL_PCC_MSG,'<PCC%>', @IMPOSTO_PCC_MSG)
            UPDATE F
              SET F.DES_MSGALIQUOTA = @MSG_PCC
              FROM #FAT_NOTAIMPRESSAO F;

             --UPDATE F
             --  SET
             --      F.DES_MSGALIQUOTA = 'PIS: '+CAST(CAST(F.VAL_PIS AS DECIMAL(18, 2)) AS VARCHAR(10))+'  COFINS: '+CAST(CAST(F.VAL_COFINS AS DECIMAL(18, 2)) AS VARCHAR(10))+'  CSLL: '+CAST(CAST(F.VAL_CSLL AS DECIMAL(18, 2)) AS VARCHAR(10))+' '+@DES_MSGComRetencao
             --FROM #FAT_NOTAIMPRESSAO F;
             --</WKF 255294 - Ricardo Gomes - Plus It>
         END;
         ELSE --IF @BO_CALCULA_IMPOST_FEDERAIS  
         BEGIN  
       
     --ZERA IMPOSTOS  
             UPDATE F
               SET
                   F.VAL_PIS = 0,
                   F.VAL_COFINS = 0,
                   F.VAL_CSLL = 0
             FROM #FAT_NOTAIMPRESSAO F;  
     
   --**** APLICA MENSAGEM PARA QUANDO NÃO TEM RETENÇÃO, EXCETO SE FOR ISENTO  
   --IF @BO_ISENTO_IMPOST_FEDERAIS=0  
   --BEGIN  
     
    --OBTEM TOTAL DE IMPOSTO FEDERAL  
             SELECT @DES_TotAliqFederal = CAST(CAST(ALIQ_PIS + ALIQ_COFINS + ALIQ_CSLL AS DECIMAL(4, 2)) AS VARCHAR)
             FROM #FAT_NOTAIMPRESSAO;  
      
     
    --APLICA MENSAGEM DIFERENTE PARA NOTAS ELETRÔNICA DE ISS E DEMAIS NOTAS  
      
    --SE FOR NOTA DE ISS  
             IF EXISTS
(
    SELECT TOP 1 *
    FROM #FAT_NOTAIMPRESSAO N
    WHERE N.TP_TRIBUTACAO = 1
)
                 BEGIN  
       
      --SE FOR FILIAL QUE EMITE NFS-E  
                     IF EXISTS
(
    SELECT TOP 1 *
    FROM FAT_CONFIGURAEXPORTANFE NFE
    WHERE NFE.COD_FILIAL = @COD_FILIAL
          AND NFE.COD_REGIONAL = @COD_REGIONAL
)
                         BEGIN  
  
        --OBTÊM MENSAGEM ESPECÍFICA PARA PREFEITURAS  
                             SET @DES_MSGSemRetencaoFinal = @DES_MSGSemRetencaoAlt;
                         END;
                         ELSE
                         BEGIN  
        
        --OBTÊM MENSAGEM GERAL  
                             SET @DES_MSGSemRetencaoFinal = @DES_MSGSemRetencao;
                         END;
                 END;
                 ELSE
                 BEGIN  
  
        --OBTÊM MENSAGEM GERAL  
                     SET @DES_MSGSemRetencaoFinal = @DES_MSGSemRetencao;
                 END;    
      
    --APLICA ATUALIZAÇÃO NA NOTA  
             UPDATE F
               SET
                   F.DES_MSGALIQUOTA = REPLACE(@DES_MSGSemRetencaoFinal, '<TOT_Aliq>', LTRIM(RTRIM(@DES_TotAliqFederal)))
             FROM #FAT_NOTAIMPRESSAO F;  
      
     
   --END --IF @BO_ISENTO_IMPOST_FEDERAIS=0  


         END; --IF @BO_CALCULA_IMPOST_FEDERAIS  
     
  
   
 --********************** Cálculo de PIS\COFINS\CSLL  
   
       
       
      
 --********************** Mensagem para BVA        


     IF EXISTS
(
    SELECT *
    FROM cor_filial C(nolock)
    WHERE C.nom_ABVD_FILIAL = 'BVA'
          AND C.cod_regional = @COD_REGIONAL
          AND C.COD_FILIAL = @COD_FILIAL
          AND Flg_SituacaoRegistro = 1
)
         BEGIN        
          
    --Produto ALUGUEL - COMPUSAFE                                  
             UPDATE N
               SET
                   N.DES_MSGNF = N.DES_MSGNF+' Recibo referente a locação do Compusafe conforme contrato'
             FROM #FAT_NOTAIMPRESSAO N
                  LEFT JOIN Fat_NotaEletronica NFE ON N.COD_FILIAL = NFE.COD_FILIAL
                                                      AND N.COD_REGIONAL = NFE.COD_REGIONAL
                                                      AND N.COD_NF = NFE.COD_NF
                                                      AND N.COD_SERIENOTA = NFE.COD_SERIENOTA
                  INNER JOIN Fat_MOVIMENTO M ON N.COD_FILIAL = M.COD_FILIAL
                                                AND N.COD_REGIONAL = M.COD_REGIONAL
                                                AND N.COD_NF = M.COD_NF
                                                AND N.COD_SERIENOTA = M.COD_SERIENOTA
             WHERE N.COD_FILIAL = @COD_FILIAL
                   AND N.COD_REGIONAL = @COD_REGIONAL
                   AND N.COD_NF = @COD_NF
                   AND N.COD_SERIENOTA = @COD_SERIENOTA
                   AND NFE.COD_REGIONAL IS NULL --EXCLUI AS ELETRÔNICAS  
                   AND M.COD_PRODUTO = 704
                   AND N.TP_SERVICO = 'L';  
  
  
    --Produto ALUGUEL - COMPUSAFE                                  
             UPDATE N
               SET
                   N.DES_MSGNF = N.DES_MSGNF+' Recibo referente a locação de equipamento de segurança eletrônica conforme contrato'
             FROM #FAT_NOTAIMPRESSAO N
                  LEFT JOIN Fat_NotaEletronica NFE ON N.COD_FILIAL = NFE.COD_FILIAL
                                                      AND N.COD_REGIONAL = NFE.COD_REGIONAL
                                                      AND N.COD_NF = NFE.COD_NF
                                                      AND N.COD_SERIENOTA = NFE.COD_SERIENOTA
                  INNER JOIN Fat_MOVIMENTO M ON N.COD_FILIAL = M.COD_FILIAL
                                                AND N.COD_REGIONAL = M.COD_REGIONAL
                                                AND N.COD_NF = M.COD_NF
                                                AND N.COD_SERIENOTA = M.COD_SERIENOTA
             WHERE N.COD_FILIAL = @COD_FILIAL
                   AND N.COD_REGIONAL = @COD_REGIONAL
                   AND N.COD_NF = @COD_NF
                   AND N.COD_SERIENOTA = @COD_SERIENOTA
                   AND NFE.COD_REGIONAL IS NULL --EXCLUI AS ELETRÔNICAS  
                   AND M.COD_PRODUTO = 711
                   AND N.TP_SERVICO = 'L';  
  
  
    --Produto ALUGUEL - COMPUSAFE                                  
             UPDATE N
               SET
                   N.DES_MSGNF = N.DES_MSGNF+' Recibo referente a locação de equipamento de contagem de numerário'
             FROM #FAT_NOTAIMPRESSAO N
                  LEFT JOIN Fat_NotaEletronica NFE ON N.COD_FILIAL = NFE.COD_FILIAL
                                                      AND N.COD_REGIONAL = NFE.COD_REGIONAL
                                                      AND N.COD_NF = NFE.COD_NF
                                                      AND N.COD_SERIENOTA = NFE.COD_SERIENOTA
                  INNER JOIN Fat_MOVIMENTO M ON N.COD_FILIAL = M.COD_FILIAL
                                                AND N.COD_REGIONAL = M.COD_REGIONAL
                                                AND N.COD_NF = M.COD_NF
                                                AND N.COD_SERIENOTA = M.COD_SERIENOTA
             WHERE N.COD_FILIAL = @COD_FILIAL
                   AND N.COD_REGIONAL = @COD_REGIONAL
                   AND N.COD_NF = @COD_NF
                   AND N.COD_SERIENOTA = @COD_SERIENOTA
                   AND NFE.COD_REGIONAL IS NULL --EXCLUI AS ELETRÔNICAS  
                   AND M.COD_PRODUTO = 748
                   AND N.TP_SERVICO = 'L';
         END;        
          
          
 --********************** Fim aplica mensagens para BVA        
             
             
 --********************** Aplicar mensagens para configurações de NFS-e         
             
 --N.DES_MSGNF            
 --VERIFICA SE A NOTA É ELETRONICA            
     IF EXISTS
(
    SELECT TOP 1 *
    FROM FAT_CONFIGURAEXPORTANFE NFE
    WHERE NFE.COD_FILIAL = @COD_FILIAL
          AND NFE.COD_REGIONAL = @COD_REGIONAL
)
         BEGIN            
             
  --VERIFICA SE O CAMPO FLG_MENSAGEMRPS TEM VALOR            
   --SE SIM, ADICIONA O COD_NF NO INICIO DA MENSAGEM            
             IF
(
    SELECT TOP 1 NFE.FLG_MENSAGEMRPS
    FROM FAT_CONFIGURAEXPORTANFE NFE
    WHERE NFE.COD_FILIAL = @COD_FILIAL
          AND NFE.COD_REGIONAL = @COD_REGIONAL
) = 1
                 BEGIN            
          
     --NOTA ELETRONICA - ATUALIZA A MENSAGEM DA NOTA COM O CODIGO NF ANTES DA MENSAGEM            
                     UPDATE N
                       SET
                           N.DES_MSGNF = '(RPS: '+CONVERT(VARCHAR(6), N.COD_NF)+')'+ISNULL(RTRIM(LTRIM(N.DES_MSGNF)), '')
                     FROM #FAT_NOTAIMPRESSAO N
                     WHERE N.COD_FILIAL = @COD_FILIAL
                           AND N.COD_REGIONAL = @COD_REGIONAL
                           AND N.COD_NF = @COD_NF
                           AND N.COD_SERIENOTA = @COD_SERIENOTA
                           AND N.TP_TRIBUTACAO = 1
                           AND N.VAL_ISS > 0;
                 END;            
              
   --VERIFICA SE O CAMPO FLG_VALORISS TEM VALOR            
   --SE SIM, ADICIONA O VALOR E ALÍQUOTA DE ISS NO FINAL DA MENSAGEM             
             IF
(
    SELECT TOP 1 NFE.FLG_VALORISS
    FROM FAT_CONFIGURAEXPORTANFE NFE
    WHERE NFE.COD_FILIAL = @COD_FILIAL
          AND NFE.COD_REGIONAL = @COD_REGIONAL
) = 1
                 BEGIN
                     UPDATE N
                       SET
                           N.DES_MSGELETRONICA = ISNULL(RTRIM(LTRIM(N.DES_MSGELETRONICA)), '')+CHAR(13)+' VALOR ISS: R$ '+REPLACE(LTRIM(STR(N.VAL_ISS, 28, 2)), '.', ',')+' ALÍQ.: '+REPLACE(LTRIM(STR(N.TX_ISS, 28, 2)), '.', ',')
                     FROM #FAT_NOTAIMPRESSAO N
                     WHERE N.COD_FILIAL = @COD_FILIAL
                           AND N.COD_REGIONAL = @COD_REGIONAL
                           AND N.COD_NF = @COD_NF
                           AND N.COD_SERIENOTA = @COD_SERIENOTA
                           AND N.TP_TRIBUTACAO = 1
                           AND N.VAL_ISS > 0;
                 END;            
          
   --VERIFICA SE UTILIZA O PARÂMETRO DE EXIBIR O SERVIÇO          
   --SE SIM, ADICIONA O ISSQN E NOME DO SERVIÇO NO CORPO DA NOTA          
             IF
(
    SELECT CONF_FIL.DES_VALORCAMPO
    FROM FAT_CONFIGURACOESFILIALNFSE CONF_FIL
         INNER JOIN FAT_CONFIGURACOESNFSE FIL ON CONF_FIL.ID_CONFIGURACOESNFSE = FIL.ID_CONFIGURACOESNFSE
    WHERE CONF_FIL.COD_REGIONAL = @COD_REGIONAL
          AND CONF_FIL.COD_FILIAL = @COD_FILIAL
          AND FIL.NOM_CAMPO = 'EXIBE SERVIÇO'
) = '1'
                 BEGIN
                     UPDATE IMP
                       SET
                           IMP.DES_MSGELETRONICA = ISNULL(RTRIM(LTRIM(IMP.DES_MSGELETRONICA)), '')+CHAR(13)+' '+FI.COD_SERVICO+CASE
                                                                                                                                   WHEN REPLACE(FI.COD_SERVICO, '.', '') = '1103'
                                                                                                                                   THEN ' - ESCOLTA '
                                                                                                                                   ELSE CASE
                                                                                                                                            WHEN REPLACE(FI.COD_SERVICO, '.', '') = '1104'
                                                                                                                                            THEN ' - PROC.NUMERÁRIO / PASTAS / TICKET / CUSTÓDIA '
                                                                                                                                            ELSE CASE
                                                                                                                                                     WHEN REPLACE(FI.COD_SERVICO, '.', '') = '1401'
                                                                                                                                                     THEN ' - ATM FLM - MANUTENÇÃO / LOCAÇÃO COMPUSAFE '
                                                                                                                                                     ELSE CASE
                                                                                                                                                              WHEN REPLACE(FI.COD_SERVICO, '.', '') = '1712'
                                                                                                                                                              THEN ' - PROC.CHEQUES / ENVELOPAMENTO / PAGAMENTO / ADM. BENS DE TERCEIROS '
                                                                                                                                                              ELSE CASE
                                                                                                                                                                       WHEN REPLACE(FI.COD_SERVICO, '.', '') = '1705'
                                                                                                                                                                       THEN ' - MÃO DE OBRA '
                                                                                                                                                                       ELSE CASE
                                                                                                                                                                                WHEN REPLACE(FI.COD_SERVICO, '.', '') = '2601'
                                                                                                                                                                                THEN ' - TRANSPORTE DE VALORES '
                                                                                                                                                                                ELSE CASE
                                                                                                                                                                                         WHEN REPLACE(FI.COD_SERVICO, '.', '') = '1102'
                                                                                                                                                                                         THEN ' - VIGILANCIA E SEGURANCA '
                                                                                                                                                                                     END
                                                                                                                                                                            END
                                                                                                                                                                   END
                                                                                                                                                          END
                                                                                                                                                 END
                                                                                                                                        END
                                                                                                                               END
                     FROM #FAT_NOTAIMPRESSAO IMP
                          INNER JOIN SIC_FAMILIAISS FI ON IMP.COD_FAMILIAISS = FI.COD_FAMILIAISS
                     WHERE IMP.COD_FILIAL = @COD_FILIAL
                           AND IMP.COD_REGIONAL = @COD_REGIONAL
                           AND IMP.COD_NF = @COD_NF
                           AND IMP.COD_SERIENOTA = @COD_SERIENOTA
                           AND IMP.TP_TRIBUTACAO = 1
                           AND IMP.VAL_ISS > 0;
                 END;            
          
          
          
          
          
   --VERIFICA SE UTILIZA O PARÂMETRO DE EXIBIR UMA MENSAGEM PARA UM SERVIÇO ESPECÍFICO          
   --SE SIM, ADICIONA A MENSAGEM NO CORPO DA NOTA        
             IF
(
    SELECT CONF_FIL.DES_VALORCAMPO
    FROM FAT_CONFIGURACOESFILIALNFSE CONF_FIL
         INNER JOIN FAT_CONFIGURACOESNFSE FIL ON CONF_FIL.ID_CONFIGURACOESNFSE = FIL.ID_CONFIGURACOESNFSE
    WHERE CONF_FIL.COD_REGIONAL = @COD_REGIONAL
          AND CONF_FIL.COD_FILIAL = @COD_FILIAL
          AND ISNULL(FIL.NOM_CAMPO, '') = 'Mensagem para Serviço'
) <> ''
                 BEGIN           
          
  --Obtêm código do Serviço          
                     SELECT @COD_SERVICO = LEFT(CONF_FIL.DES_VALORCAMPO, CHARINDEX('|', CONF_FIL.DES_VALORCAMPO)-1)
                     FROM #FAT_NOTAIMPRESSAO IMP
                          INNER JOIN SIC_FAMILIAISS FI ON IMP.COD_FAMILIAISS = FI.COD_FAMILIAISS
                          INNER JOIN FAT_CONFIGURACOESFILIALNFSE CONF_FIL ON CONF_FIL.COD_REGIONAL = IMP.COD_REGIONAL
                                                                             AND CONF_FIL.COD_FILIAL = IMP.COD_FILIAL
                          INNER JOIN FAT_CONFIGURACOESNFSE FIL ON CONF_FIL.ID_CONFIGURACOESNFSE = FIL.ID_CONFIGURACOESNFSE
                     WHERE IMP.COD_FILIAL = @COD_FILIAL
                           AND IMP.COD_REGIONAL = @COD_REGIONAL
                           AND IMP.COD_NF = @COD_NF
                           AND IMP.COD_SERIENOTA = @COD_SERIENOTA
                           AND IMP.TP_TRIBUTACAO = 1
                           AND IMP.VAL_ISS > 0
                           AND FIL.NOM_CAMPO = 'Mensagem para Serviço';          
          
          
  --IDENTIFICA O SERVIÇO DA MENSAGEM E ATRIBUI A DEVIDA MENSAGEM          
                     UPDATE IMP
                       SET
                           IMP.DES_MSGELETRONICA = ISNULL(RTRIM(LTRIM(IMP.DES_MSGELETRONICA)), '')+' '+RIGHT(CONF_FIL.DES_VALORCAMPO, LEN(CONF_FIL.DES_VALORCAMPO)-CHARINDEX('|', CONF_FIL.DES_VALORCAMPO))
                     FROM #FAT_NOTAIMPRESSAO IMP
                          INNER JOIN SIC_FAMILIAISS FI ON IMP.COD_FAMILIAISS = FI.COD_FAMILIAISS
                          INNER JOIN FAT_CONFIGURACOESFILIALNFSE CONF_FIL ON CONF_FIL.COD_REGIONAL = IMP.COD_REGIONAL
                                                                             AND CONF_FIL.COD_FILIAL = IMP.COD_FILIAL
                          INNER JOIN FAT_CONFIGURACOESNFSE FIL ON CONF_FIL.ID_CONFIGURACOESNFSE = FIL.ID_CONFIGURACOESNFSE
                     WHERE IMP.COD_FILIAL = @COD_FILIAL
                           AND IMP.COD_REGIONAL = @COD_REGIONAL
                           AND IMP.COD_NF = @COD_NF
                           AND IMP.COD_SERIENOTA = @COD_SERIENOTA
                           AND IMP.TP_TRIBUTACAO = 1
                           AND IMP.VAL_ISS > 0
                           AND FIL.NOM_CAMPO = 'Mensagem para Serviço'
                           AND --PARÂMETRO DO SERVIÇO          
                           LTRIM(RTRIM(FI.COD_SERVICO)) = @COD_SERVICO;
                 END;          
          
          
          
   --VERIFICA SE UTILIZA O PARÂMETRO DE EXIBIR A CIDADE          
   --SE SIM, ADICIONA O NOME DA CIDADE NO CORPO DA NOTA          
             IF
(
    SELECT CONF_FIL.DES_VALORCAMPO
    FROM FAT_CONFIGURACOESFILIALNFSE CONF_FIL
         INNER JOIN FAT_CONFIGURACOESNFSE FIL ON CONF_FIL.ID_CONFIGURACOESNFSE = FIL.ID_CONFIGURACOESNFSE
    WHERE CONF_FIL.COD_REGIONAL = @COD_REGIONAL
          AND CONF_FIL.COD_FILIAL = @COD_FILIAL
          AND FIL.NOM_CAMPO = 'EXIBE CIDADE'
) = '1'
                 BEGIN
                     UPDATE N
                       SET
                           N.DES_MSGELETRONICA = ISNULL(RTRIM(LTRIM(N.DES_MSGELETRONICA)), '')+CHAR(13)+' SERVIÇO REALIZADO EM: '+CASE
                                                                                                                                      WHEN N.COD_FAMILIAISS <> 7
                                                                                                                                      THEN CT.NOM_CIDADE
                                                                                                                                      ELSE CI.NOM_CIDADE
                                                                                                                                  END
                     FROM #FAT_NOTAIMPRESSAO N
                          INNER JOIN FAT_MOVIMENTO M ON M.COD_REGIONAL = N.COD_REGIONAL
                                                        AND M.COD_FILIAL = N.COD_FILIAL
                                                        AND M.COD_NF = N.COD_NF
                                                        AND M.COD_SERIENOTA = N.COD_SERIENOTA
                          INNER JOIN SIC_TRIBUTACAO TR ON N.COD_TRIBUTACAO = TR.COD_TRIBUTACAO
                          INNER JOIN COR_CIDADE CT ON TR.COD_CIDADE = CT.COD_CIDADE
                          INNER JOIN SIC_SUBCLIENTE SC ON SC.CGC_SUBCLIENTE = M.CGC_ORIGEM
                                                          AND SC.SEQ = M.SEQ_ORIGEM
                          INNER JOIN COR_CIDADE CI ON CI.COD_CIDADE = SC.COD_CIDADE
                     WHERE N.COD_FILIAL = @COD_FILIAL
                           AND N.COD_REGIONAL = @COD_REGIONAL
                           AND N.COD_NF = @COD_NF
                           AND N.COD_SERIENOTA = @COD_SERIENOTA
                           AND N.TP_TRIBUTACAO = 1
                           AND N.VAL_ISS > 0;
                 END;
         END;           
         
         
 --********************** Fim para aplicar mensagens para configurações de NFS-e         

     IF NOT EXISTS
(
    SELECT *
    FROM cor_filial C(nolock)
    WHERE C.nom_ABVD_FILIAL = 'BVA'
          AND C.cod_regional = @COD_REGIONAL
          AND C.COD_FILIAL = @COD_FILIAL
)
         BEGIN        
    --DESCONSIDERA EMPRESA BVA        
            
    --NOTA NÃO ELETRONICA - ADICIONA MENSAGEM DE NÃO SUBSTITUTO           


             UPDATE N
               SET
                   N.DES_MSGNF = LEFT(RTRIM(N.DES_MSGNF)+' - ISS RECOLHIDO PELO PRESTADOR', 200)
             FROM #FAT_NOTAIMPRESSAO N
             WHERE N.IC_SUBSTITUICAOTRIBUTARIAISS = 'N'
                   AND N.VAL_ISS > 0
                   AND N.COD_REGIONAL = @COD_REGIONAL
                   AND N.COD_FILIAL = @COD_FILIAL
                   AND N.COD_NF = @COD_NF
                   AND N.COD_SERIENOTA = @COD_SERIENOTA
                   AND N.Flg_IsencaoISS = 0;
         END;            
          
--ATUALIZA A DESCRIÇÃO DA NOTA NA TABELA FAT_NOTA, IGUAL A FAT_NOTAIMPRESSAO          
     UPDATE N
       SET
           N.DES_MSGNF = N2.DES_MSGNF
     FROM FAT_NOTA N
          INNER JOIN #FAT_NOTAIMPRESSAO N2 ON N.COD_REGIONAL = N2.COD_REGIONAL
                                              AND N.COD_FILIAL = N2.COD_FILIAL
                                              AND N.COD_NF = N2.COD_NF
                                              AND N.COD_SERIENOTA = N2.COD_SERIENOTA
     WHERE N.COD_REGIONAL = @COD_REGIONAL
           AND N.COD_FILIAL = @COD_FILIAL
           AND N.COD_NF = @COD_NF
           AND N.COD_SERIENOTA = @COD_SERIENOTA;            
          
            
 --ATUALIZA ENDEREÇO DE COBRANÇA CONFORME SIC_SUBCLIENTE              
     UPDATE N
       SET
           N.END_CLIENTE = SC.END_SUBCLIENTE+' '+ISNULL(SC.END_NUMERO, ''),
           N.CIDADE_CLIENTE = SC.NOM_CIDADE,
           N.CEP_CLIENTE = SC.CEP_SUBCLIENTE,
           N.NOM_UF = SC.UF_SUBCLIENTE,
           N.CGC_CLIENTE = DBO.FN_FAT_FORMATACGC(SC.COD_CNPJCPF),
           N.INSCE_CLIENTE = SC.COD_INSCE,
           N.INSCM_CLIENTE = SC.COD_INSCM
     FROM #FAT_NOTAIMPRESSAO N
          INNER JOIN FAT_NOTA N2 ON N2.COD_FILIAL = N.COD_FILIAL
                                    AND N2.COD_REGIONAL = N.COD_REGIONAL
                                    AND N2.COD_NF = N.COD_NF
                                    AND N2.COD_SERIENOTA = N.COD_SERIENOTA
          INNER JOIN SIC_CLIENTETRIBUTACAO T ON N2.CGC_CLIENTE = T.CGC_CLIENTE
                                                AND N2.SEQ_CLIENTE = T.SEQ_CLIENTE
                                                AND N2.COD_CLIENTETRIBUTACAO = T.COD_CLIENTETRIBUTACAO
          INNER JOIN SIC_SUBCLIENTE SC ON T.CGC_SUBCLIENTESACADO = SC.CGC_SUBCLIENTE
                                          AND T.SEQ_SUBCLIENTESACADO = SC.SEQ
     WHERE N.FLG_DADOSSUBCLIENTESACADO = 'S'
           AND N.COD_FILIAL = @COD_FILIAL
           AND N.COD_REGIONAL = @COD_REGIONAL
           AND N.COD_NF = @COD_NF
           AND N.COD_SERIENOTA = @COD_SERIENOTA;
     INSERT INTO #DADOS
(COD_REGIONAL,
 COD_FILIAL,
 COD_NF,
 QTDE,
 UNIDADE,
 DES_PRODUTO,
 COD_TRIBUTACAO,
 VAL_BRUTO,
 COD_SERIENOTA
)
            SELECT N.COD_REGIONAL,
                   N.COD_FILIAL,
                   N.COD_NF,
                   SUM(CASE
                           WHEN P.COD_FAMILIA = 2
                           THEN QTD_PROCESSAMENTO
                           ELSE 0
                       END) QTDE,
                   '' UNIDADE,
                   P.DES_PRODUTO,
                   M.COD_TRIBUTACAO,
                   SUM(M.VAL_BRUTO) VAL_BRUTO,
                   N.COD_SERIENOTA
            FROM #FAT_NOTAIMPRESSAO N
                 INNER JOIN FAT_MOVIMENTO M ON N.COD_NF = M.COD_NF
                                               AND N.COD_SERIENOTA = M.COD_SERIENOTA
                                               AND N.COD_FILIAL = M.COD_FILIAL
                                               AND N.COD_REGIONAL = M.COD_REGIONAL
                 INNER JOIN SIC_PRODUTO P ON M.COD_PRODUTO = P.COD_PRODUTO
            WHERE M.VAL_BRUTO <> 0
                  AND N.COD_FILIAL = @COD_FILIAL
                  AND N.COD_REGIONAL = @COD_REGIONAL
                  AND N.COD_NF = @COD_NF
            GROUP BY N.COD_REGIONAL,
                     N.COD_FILIAL,
                     N.COD_NF,
                     M.COD_TRIBUTACAO,
                     DES_PRODUTO,
                     N.COD_SERIENOTA;            
            
            
 --ATUALIZA MENSAGEM DE SUBSTITUIÇÃO_ICMS CONFORME MUNICIPIO.----------------------------            
     UPDATE N
       SET
           N.DES_MENSAGEMSUBS_ICMS = TR.DES_SUBSTRIBICMS
     FROM #FAT_NOTAIMPRESSAO N
          INNER JOIN SIC_TRIBUTACAO TR ON TR.COD_TRIBUTACAO = N.COD_TRIBUTACAO
     WHERE N.TP_TRIBUTACAO >= 2;              
            
            
 -- TRATA QUANDO EXISTIREM MAIS DE 6 PRODUTOS DE ISS              
     SELECT COD_REGIONAL,
            COD_FILIAL,
            COD_NF,
            COD_TRIBUTACAO,
            COUNT(*) CONTADOR,
            MIN(COD_SEQ) SEQ_MIN,
            MAX(COD_SEQ) SEQ_MAX,
            COD_SERIENOTA
     INTO #QTD_ISS
     FROM #DADOS
     WHERE COD_TRIBUTACAO = 1
     GROUP BY COD_REGIONAL,
              COD_FILIAL,
              COD_NF,
              COD_TRIBUTACAO,
              COD_SERIENOTA
     HAVING COUNT(*) > 6;              
            
  --ALIMENTAÇÃO DA TABELA #MAIS_ISS            
     SELECT Q.COD_REGIONAL,
            Q.COD_FILIAL,
            Q.COD_NF,
            SUM(QTDE) QTDE,
            MAX(UNIDADE) UNIDADE,
            'OUTROS' DES_PRODUTO,
            Q.COD_TRIBUTACAO,
            SUM(VAL_BRUTO) VAL_BRUTO,
            Q.COD_SERIENOTA
     INTO #MAIS_ISS
     FROM #DADOS D
          INNER JOIN #QTD_ISS Q ON D.COD_REGIONAL = Q.COD_REGIONAL
                                   AND D.COD_FILIAL = Q.COD_FILIAL
                                   AND D.COD_NF = Q.COD_NF
                                   AND D.COD_SERIENOTA = Q.COD_SERIENOTA
                                   AND D.COD_TRIBUTACAO = Q.COD_TRIBUTACAO
     WHERE Q.COD_TRIBUTACAO = 1
           AND COD_SEQ BETWEEN SEQ_MIN + 5 AND SEQ_MAX
     GROUP BY Q.COD_REGIONAL,
              Q.COD_FILIAL,
              Q.COD_NF,
              Q.COD_TRIBUTACAO,
              Q.COD_SERIENOTA;
     DELETE D
     FROM #DADOS D
          INNER JOIN #QTD_ISS Q ON D.COD_REGIONAL = Q.COD_REGIONAL
                                   AND D.COD_FILIAL = Q.COD_FILIAL
                                   AND D.COD_NF = Q.COD_NF
                                   AND D.COD_SERIENOTA = Q.COD_SERIENOTA
                                   AND D.COD_TRIBUTACAO = Q.COD_TRIBUTACAO
     WHERE Q.COD_TRIBUTACAO = 1
           AND COD_SEQ BETWEEN SEQ_MIN + 5 AND SEQ_MAX;
     INSERT INTO #DADOS
            SELECT *
            FROM #MAIS_ISS;              
            
            
 -- TRATA QUANDO EXISTIREM MAIS DE 4 PRODUTOS DE ICMS              
     SELECT COD_REGIONAL,
            COD_FILIAL,
            COD_NF,
            COD_TRIBUTACAO,
            COUNT(*) CONTADOR,
            MIN(COD_SEQ) SEQ_MIN,
            MAX(COD_SEQ) SEQ_MAX,
            COD_SERIENOTA
     INTO #QTD_ICMS
     FROM #DADOS
     WHERE COD_TRIBUTACAO IN(2, 3, 5)
     GROUP BY COD_REGIONAL,
              COD_FILIAL,
              COD_NF,
              COD_TRIBUTACAO,
              COD_SERIENOTA
     HAVING COUNT(*) > 4;
     SELECT Q.COD_REGIONAL,
            Q.COD_FILIAL,
            Q.COD_NF,
            SUM(QTDE) QTDE,
            MAX(UNIDADE) UNIDADE,
            'OUTROS' DES_PRODUTO,
            Q.COD_TRIBUTACAO,
            SUM(VAL_BRUTO) VAL_BRUTO,
            Q.COD_SERIENOTA
     INTO #MAIS_ICMS
     FROM #DADOS D
          INNER JOIN #QTD_ICMS Q ON D.COD_REGIONAL = Q.COD_REGIONAL
                                    AND D.COD_FILIAL = Q.COD_FILIAL
                                    AND D.COD_NF = Q.COD_NF
                                    AND D.COD_SERIENOTA = Q.COD_SERIENOTA
                                    AND D.COD_TRIBUTACAO = Q.COD_TRIBUTACAO
     WHERE Q.COD_TRIBUTACAO IN(2, 3, 5)
          AND COD_SEQ BETWEEN SEQ_MIN + 3 AND SEQ_MAX
     GROUP BY Q.COD_REGIONAL,
              Q.COD_FILIAL,
              Q.COD_NF,
              Q.COD_TRIBUTACAO,
              Q.COD_SERIENOTA;
     DELETE D
     FROM #DADOS D
          INNER JOIN #QTD_ICMS Q ON D.COD_REGIONAL = Q.COD_REGIONAL
                                    AND D.COD_FILIAL = Q.COD_FILIAL
                                    AND D.COD_NF = Q.COD_NF
                                    AND D.COD_SERIENOTA = Q.COD_SERIENOTA
                                    AND D.COD_TRIBUTACAO = Q.COD_TRIBUTACAO
     WHERE Q.COD_TRIBUTACAO IN(2, 3, 5)
           AND COD_SEQ BETWEEN SEQ_MIN + 3 AND SEQ_MAX;
     INSERT INTO #DADOS
            SELECT *
            FROM #MAIS_ICMS;              
            
            
 -- COLOCA UNIDADE              
     UPDATE #DADOS
       SET
           UNIDADE = '.';
     UPDATE #DADOS
       SET
           UNIDADE = 'MIL.'
     WHERE QTDE > 0;            
 ---------------------------------------------------------------------------------------            

     SET @COD_ISS =
(
    SELECT CASE
               WHEN N.TP_TRIBUTACAO = 1
               THEN            
  
/*    CASE N.TP_SERVICO     SDM 11673 INICIO    
      WHEN 'E' THEN            
       '11.03'            
      WHEN 'N' THEN            
       '11.04'            
      WHEN 'F' THEN            
       '14.01'            
      WHEN 'C' THEN            
       '17.12'            
      WHEN 'P' THEN            
       '17.12'            
      WHEN 'M' THEN            
       '17.05'            
      WHEN 'T' THEN            
       '26.01'            
      WHEN 'V' THEN            
       '11.02'                                                                           
     END            
 */        
        
        
        
           FI.COD_SERVICO
           END
    FROM #FAT_NOTAIMPRESSAO N
         INNER JOIN SIC_FAMILIAISS FI ON N.COD_FAMILIAISS = FI.COD_FAMILIAISS
    WHERE N.COD_FILIAL = @COD_FILIAL
          AND N.COD_REGIONAL = @COD_REGIONAL
          AND N.COD_NF = @COD_NF
          AND N.COD_SERIENOTA = @COD_SERIENOTA
);
     DECLARE @PRODUTO VARCHAR(70), @TRIBUTACAO INT;
     SET @CONTADOR = 1;
     WHILE
(
    SELECT COUNT(*)
    FROM #DADOS
) > 0
         BEGIN            
  --SET ROWCOUNT 0            

             SET @PRODUTO =
(
    SELECT TOP 1 DES_PRODUTO
    FROM #DADOS
);
             SET @TRIBUTACAO =
(
    SELECT TOP 1 COD_TRIBUTACAO
    FROM #DADOS
);
             IF @TRIBUTACAO = 1
                 BEGIN
                     IF @CONTADOR = 1
                         BEGIN
                             UPDATE N
                               SET
                                   N.DES_QTD_ISS = CASE
                                                       WHEN ISNULL(N.DES_QTD_ISS, '') = ''
                                                       THEN RTRIM(LTRIM(CASE
                                                                            WHEN D.QTDE = 0
                                                                            THEN CHAR(13)
                                                                            ELSE STR(D.QTDE, 9, 2)
                                                                        END))
                                                       ELSE RTRIM(LTRIM(ISNULL(N.DES_QTD_ISS, '')))+RTRIM(LTRIM(CASE
                                                                                                                    WHEN D.QTDE = 0
                                                                                                                    THEN ''
                                                                                                                    ELSE STR(D.QTDE, 9, 2)
                                                                                                                END))+CHAR(13)
                                                   END,
                                   N.DES_UNID_ISS = CASE
                                                        WHEN ISNULL(N.DES_UNID_ISS, '') = ''
                                                        THEN RTRIM(LTRIM(CASE
                                                                             WHEN D.UNIDADE = '.'
                                                                             THEN CHAR(13)
                                                                             ELSE D.UNIDADE+CHAR(13)
                                                                         END))
                                                        ELSE RTRIM(LTRIM(ISNULL(N.DES_UNID_ISS, '')))+RTRIM(LTRIM(CASE
                                                                                                                      WHEN D.UNIDADE = '.'
                                                                                                                      THEN ''
                                                                                                                      ELSE D.UNIDADE
                                                                                                                  END))+CHAR(13)
                                                    END,
                                   N.DES_PRODUTO_ISS = CONVERT(CHAR(40), D.DES_PRODUTO),
                                   N.DES_PRODUTO_ISS1 = CASE
                                                            WHEN LEN(ISNULL(N.DES_PRODUTO_ISS1, '')) = 0
                                                            THEN CASE
                                                                     WHEN SUBSTRING(D.DES_PRODUTO, 1, 40) <> ''
                                                                     THEN CASE
                                                                              WHEN LTRIM(RTRIM(@COD_ISS)) <> ''
                                                                              THEN LTRIM(RTRIM(@COD_ISS))+'-'
                                                                              ELSE ''
                                                                          END+SUBSTRING(D.DES_PRODUTO, 1, 40)
                                                                     ELSE ''
                                                                 END
                                                            ELSE ISNULL(N.DES_PRODUTO_ISS1, '')
                                                        END,
                                   N.DES_VALOR_ISS = CASE
                                                         WHEN ISNULL(N.DES_VALOR_ISS, '') = ''
                                                         THEN RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                         ELSE RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+CHAR(13)+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                     END,
                                   N.COD_ISS = @COD_ISS
                             FROM #FAT_NOTAIMPRESSAO N
                                  INNER JOIN #DADOS D ON N.COD_NF = D.COD_NF
                                                         AND N.COD_SERIENOTA = D.COD_SERIENOTA
                                                         AND N.COD_FILIAL = D.COD_FILIAL
                                                         AND N.COD_REGIONAL = D.COD_REGIONAL
                             WHERE D.COD_TRIBUTACAO = @TRIBUTACAO
                                   AND D.DES_PRODUTO = @PRODUTO
                                   AND N.COD_FILIAL = @COD_FILIAL
                                   AND N.COD_REGIONAL = @COD_REGIONAL
                                   AND N.COD_NF = @COD_NF
                                   AND N.COD_SERIENOTA = @COD_SERIENOTA;
                         END;
                     IF @CONTADOR = 2
                         BEGIN
                             UPDATE N
                               SET
                                   N.DES_PRODUTO_ISS2 = CASE
                                                            WHEN LEN(ISNULL(N.DES_PRODUTO_ISS2, '')) = 0
                                                            THEN CASE
                                                                     WHEN SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40) <> ''
                                                                     THEN CASE
                                                                              WHEN LTRIM(RTRIM(@COD_ISS)) <> ''
                                                                              THEN LTRIM(RTRIM(@COD_ISS))+'-'
                                                                              ELSE ''
                                                                          END+SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40)
                                                                     ELSE ''
                                                                 END
                                                            ELSE ISNULL(N.DES_PRODUTO_ISS2, '')
                                                        END,
                                   N.DES_VALOR_ISS = CASE
                                                         WHEN ISNULL(N.DES_VALOR_ISS, '') = ''
                                                         THEN RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                         ELSE RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+CHAR(13)+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                     END
                             FROM #FAT_NOTAIMPRESSAO N
                                  INNER JOIN #DADOS D ON N.COD_NF = D.COD_NF
                                                         AND N.COD_SERIENOTA = D.COD_SERIENOTA
                                                         AND N.COD_FILIAL = D.COD_FILIAL
                                                         AND N.COD_REGIONAL = D.COD_REGIONAL
                             WHERE D.COD_TRIBUTACAO = @TRIBUTACAO
                                   AND D.DES_PRODUTO = @PRODUTO
                                   AND N.COD_FILIAL = @COD_FILIAL
                                   AND N.COD_REGIONAL = @COD_REGIONAL
                                   AND N.COD_NF = @COD_NF
                                   AND N.COD_SERIENOTA = @COD_SERIENOTA;
                         END;
                         ELSE
                     IF @CONTADOR = 3
                         BEGIN
                             UPDATE N
                               SET
                                   N.DES_PRODUTO_ISS3 = CASE
                                                            WHEN LEN(ISNULL(N.DES_PRODUTO_ISS3, '')) = 0
                                                            THEN CASE
                                                                     WHEN SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40) <> ''
                                                                     THEN CASE
                                                                              WHEN LTRIM(RTRIM(@COD_ISS)) <> ''
                                                                              THEN LTRIM(RTRIM(@COD_ISS))+'-'
                                                                              ELSE ''
                                                                          END+SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40)
                                                                     ELSE ''
                                                                 END
                                                            ELSE ISNULL(N.DES_PRODUTO_ISS3, '')
                                                        END,
                                   N.DES_VALOR_ISS = CASE
                                                         WHEN ISNULL(N.DES_VALOR_ISS, '') = ''
                                                         THEN RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                         ELSE RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+CHAR(13)+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                     END
                             FROM #FAT_NOTAIMPRESSAO N
                                  INNER JOIN #DADOS D ON N.COD_NF = D.COD_NF
                                                         AND N.COD_SERIENOTA = D.COD_SERIENOTA
                                                         AND N.COD_FILIAL = D.COD_FILIAL
                                                         AND N.COD_REGIONAL = D.COD_REGIONAL
                             WHERE D.COD_TRIBUTACAO = @TRIBUTACAO
                                   AND D.DES_PRODUTO = @PRODUTO
                                   AND N.COD_FILIAL = @COD_FILIAL
                                   AND N.COD_REGIONAL = @COD_REGIONAL
                                   AND N.COD_NF = @COD_NF
                                   AND N.COD_SERIENOTA = @COD_SERIENOTA;
                         END;
                         ELSE
                     IF @CONTADOR = 4
                         BEGIN
                             UPDATE N
                               SET
                                   N.DES_PRODUTO_ISS4 = CASE
                                                            WHEN LEN(ISNULL(N.DES_PRODUTO_ISS4, '')) = 0
                                                            THEN CASE
                                                                     WHEN SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40) <> ''
                                                                     THEN CASE
                                                                              WHEN LTRIM(RTRIM(@COD_ISS)) <> ''
                                                                              THEN LTRIM(RTRIM(@COD_ISS))+'-'
                                                                              ELSE ''
                                                                          END+SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40)
                                                                     ELSE ''
                                                                 END
                                                            ELSE ISNULL(N.DES_PRODUTO_ISS4, '')
                                                        END,
                                   N.DES_VALOR_ISS = CASE
                                                         WHEN ISNULL(N.DES_VALOR_ISS, '') = ''
                                                         THEN RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                         ELSE RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+CHAR(13)+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                     END
                             FROM #FAT_NOTAIMPRESSAO N
                                  INNER JOIN #DADOS D ON N.COD_NF = D.COD_NF
                                                         AND N.COD_SERIENOTA = D.COD_SERIENOTA
                                                         AND N.COD_FILIAL = D.COD_FILIAL
                                                         AND N.COD_REGIONAL = D.COD_REGIONAL
                             WHERE D.COD_TRIBUTACAO = @TRIBUTACAO
                                   AND D.DES_PRODUTO = @PRODUTO
                                   AND N.COD_FILIAL = @COD_FILIAL
                                   AND N.COD_REGIONAL = @COD_REGIONAL
                                   AND N.COD_NF = @COD_NF
                                   AND N.COD_SERIENOTA = @COD_SERIENOTA;
                         END;
                         ELSE
                     IF @CONTADOR = 5
                         BEGIN
                             UPDATE N
                               SET
                                   N.DES_PRODUTO_ISS5 = CASE
                                                            WHEN LEN(ISNULL(N.DES_PRODUTO_ISS5, '')) = 0
                                                            THEN CASE
                                                                     WHEN SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40) <> ''
                                                                     THEN CASE
                                                                              WHEN LTRIM(RTRIM(@COD_ISS)) <> ''
                                                                              THEN LTRIM(RTRIM(@COD_ISS))+'-'
                                                                              ELSE ''
                                                                          END+SUBSTRING(D.DES_PRODUTO, 1, 40)
                                                                     ELSE ''
                                                                 END
                                                            ELSE ISNULL(N.DES_PRODUTO_ISS5, '')
                                                        END,
                                   N.DES_VALOR_ISS = CASE
                                                         WHEN ISNULL(N.DES_VALOR_ISS, '') = ''
                                                         THEN RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                         ELSE RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+CHAR(13)+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                     END
                             FROM #FAT_NOTAIMPRESSAO N
                                  INNER JOIN #DADOS D ON N.COD_NF = D.COD_NF
                                                         AND N.COD_SERIENOTA = D.COD_SERIENOTA
                                                         AND N.COD_FILIAL = D.COD_FILIAL
                                                         AND N.COD_REGIONAL = D.COD_REGIONAL
                             WHERE D.COD_TRIBUTACAO = @TRIBUTACAO
                                   AND D.DES_PRODUTO = @PRODUTO
                                   AND N.COD_FILIAL = @COD_FILIAL
                                   AND N.COD_REGIONAL = @COD_REGIONAL
                                   AND N.COD_NF = @COD_NF
                                   AND N.COD_SERIENOTA = @COD_SERIENOTA;
                         END;
                         ELSE
                     IF @CONTADOR = 6
                         BEGIN
                             UPDATE N
                               SET
                                   N.DES_PRODUTO_ISS6 = CASE
                                                            WHEN LEN(ISNULL(N.DES_PRODUTO_ISS6, '')) = 0
                                                            THEN CASE
                                                                     WHEN SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40) <> ''
                                                                     THEN CASE
                                                                              WHEN LTRIM(RTRIM(@COD_ISS)) <> ''
                                                                              THEN LTRIM(RTRIM(@COD_ISS))+'-'
                                                                              ELSE ''
                                                                          END+SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40)
                                                                     ELSE ''
                                                                 END
                                                            ELSE ISNULL(N.DES_PRODUTO_ISS6, '')
                                                        END,
                                   N.DES_VALOR_ISS = CASE
                                                         WHEN ISNULL(N.DES_VALOR_ISS, '') = ''
                                                         THEN RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                         ELSE RTRIM(LTRIM(ISNULL(N.DES_VALOR_ISS, '')))+CHAR(13)+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                     END
                             FROM #FAT_NOTAIMPRESSAO N
                                  INNER JOIN #DADOS D ON N.COD_NF = D.COD_NF
                                                         AND N.COD_SERIENOTA = D.COD_SERIENOTA
                                                         AND N.COD_FILIAL = D.COD_FILIAL
                                                         AND N.COD_REGIONAL = D.COD_REGIONAL
                             WHERE D.COD_TRIBUTACAO = @TRIBUTACAO
                                   AND D.DES_PRODUTO = @PRODUTO
                                   AND N.COD_FILIAL = @COD_FILIAL
                                   AND N.COD_REGIONAL = @COD_REGIONAL
                                   AND N.COD_NF = @COD_NF
                                   AND N.COD_SERIENOTA = @COD_SERIENOTA;
                         END;
                 END;
                 ELSE
                 BEGIN
                     IF @CONTADOR = 1
                         BEGIN
                             UPDATE N
                               SET
                                   N.DES_QTD_ICMS = CASE
                                                        WHEN ISNULL(N.DES_QTD_ICMS, '') = ''
                                                        THEN RTRIM(LTRIM(CASE
                                                                             WHEN D.QTDE = 0
                                                                             THEN CHAR(13)
                                                                             ELSE STR(D.QTDE, 9, 2)
                                                                         END))
                                                        ELSE RTRIM(LTRIM(ISNULL(N.DES_QTD_ICMS, 0)))+RTRIM(LTRIM(CASE
                                                                                                                     WHEN D.QTDE = 0
                                                                                                                     THEN ''
                                                                                                                     ELSE STR(D.QTDE, 9, 2)
                                                                                                                 END))+CHAR(13)
                                                    END,
                                   N.DES_UNID_ICMS = CASE
                                                         WHEN ISNULL(N.DES_UNID_ICMS, '') = ''
                                                         THEN RTRIM(LTRIM(CASE
                                                                              WHEN D.UNIDADE = '.'
                                                                              THEN CHAR(13)
                                                                              ELSE D.UNIDADE+CHAR(13)
                                                                          END))
                                                         ELSE RTRIM(LTRIM(ISNULL(N.DES_UNID_ICMS, '')))+RTRIM(LTRIM(CASE
                                                                                                                        WHEN D.UNIDADE = '.'
                                                                                                                        THEN ''
                                                                                                                        ELSE D.UNIDADE
                                                                                                                    END))+CHAR(13)
                                                     END,
                                   N.DES_PRODUTO_ICMS = CONVERT(CHAR(40), ISNULL(D.DES_PRODUTO, '')),
                                   N.DES_PRODUTO_ICMS1 = CASE
                                                             WHEN LEN(ISNULL(N.DES_PRODUTO_ICMS1, '')) = 0
                                                             THEN SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40)
                                                             ELSE ISNULL(N.DES_PRODUTO_ICMS1, '')
                                                         END,
                                   N.DES_VALOR_ICMS = CASE
                                                          WHEN ISNULL(N.DES_VALOR_ICMS, '') = ''
                                                          THEN LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                          ELSE RTRIM(LTRIM(ISNULL(N.DES_VALOR_ICMS, '')))+CHAR(13)+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                      END
                             FROM #FAT_NOTAIMPRESSAO N
                                  INNER JOIN #DADOS D ON N.COD_NF = D.COD_NF
                                                         AND N.COD_SERIENOTA = D.COD_SERIENOTA
                                                         AND N.COD_FILIAL = D.COD_FILIAL
                                                         AND N.COD_REGIONAL = D.COD_REGIONAL
                             WHERE D.COD_TRIBUTACAO = @TRIBUTACAO
                                   AND D.DES_PRODUTO = @PRODUTO
                                   AND N.COD_FILIAL = @COD_FILIAL
                                   AND N.COD_REGIONAL = @COD_REGIONAL
                                   AND N.COD_NF = @COD_NF
                                   AND N.COD_SERIENOTA = @COD_SERIENOTA;
                         END;
                         ELSE
                     IF @CONTADOR = 2
                         BEGIN
                             UPDATE N
                               SET
                                   N.DES_PRODUTO_ICMS2 = CASE
                                                             WHEN LEN(ISNULL(N.DES_PRODUTO_ICMS2, '')) = 0
                                                             THEN SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40)
                                                             ELSE ISNULL(N.DES_PRODUTO_ICMS2, '')
                                                         END,
                                   N.DES_VALOR_ICMS = CASE
                                                          WHEN ISNULL(N.DES_VALOR_ICMS, '') = ''
                                                          THEN LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                          ELSE RTRIM(LTRIM(ISNULL(N.DES_VALOR_ICMS, '')))+CHAR(13)+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                      END
                             FROM #FAT_NOTAIMPRESSAO N
                                  INNER JOIN #DADOS D ON N.COD_NF = D.COD_NF
                                                         AND N.COD_SERIENOTA = D.COD_SERIENOTA
                                                         AND N.COD_FILIAL = D.COD_FILIAL
                                                         AND N.COD_REGIONAL = D.COD_REGIONAL
                             WHERE D.COD_TRIBUTACAO = @TRIBUTACAO
                                   AND D.DES_PRODUTO = @PRODUTO
                                   AND N.COD_FILIAL = @COD_FILIAL
                                   AND N.COD_REGIONAL = @COD_REGIONAL
                                   AND N.COD_NF = @COD_NF
                                   AND N.COD_SERIENOTA = @COD_SERIENOTA;
                         END;
                         ELSE
                     IF @CONTADOR = 3
                         BEGIN
                             UPDATE N
                               SET
                                   N.DES_PRODUTO_ICMS3 = CASE
                                                             WHEN LEN(ISNULL(N.DES_PRODUTO_ICMS3, '')) = 0
                                                             THEN SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40)
                                                             ELSE ISNULL(N.DES_PRODUTO_ICMS3, '')
                                                         END,
                                   N.DES_VALOR_ICMS = CASE
                                                          WHEN ISNULL(N.DES_VALOR_ICMS, '') = ''
                                                          THEN LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                          ELSE RTRIM(LTRIM(ISNULL(N.DES_VALOR_ICMS, '')))+CHAR(13)+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                      END
                             FROM #FAT_NOTAIMPRESSAO N
                                  INNER JOIN #DADOS D ON N.COD_NF = D.COD_NF
                                                         AND N.COD_SERIENOTA = D.COD_SERIENOTA
                                                         AND N.COD_FILIAL = D.COD_FILIAL
                                                         AND N.COD_REGIONAL = D.COD_REGIONAL
                             WHERE D.COD_TRIBUTACAO = @TRIBUTACAO
                                   AND D.DES_PRODUTO = @PRODUTO
                                   AND N.COD_FILIAL = @COD_FILIAL
                                   AND N.COD_REGIONAL = @COD_REGIONAL
                                   AND N.COD_NF = @COD_NF
                                   AND N.COD_SERIENOTA = @COD_SERIENOTA;
                         END;
                         ELSE
                     IF @CONTADOR = 4
                         BEGIN
                             UPDATE N
                               SET
                                   N.DES_PRODUTO_ICMS4 = CASE
                                                             WHEN LEN(ISNULL(N.DES_PRODUTO_ICMS4, '')) = 0
                                                             THEN SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 40)
                                                             ELSE ISNULL(N.DES_PRODUTO_ICMS4, '')
                                                         END,
                                   N.DES_VALOR_ICMS = CASE
                                                          WHEN ISNULL(N.DES_VALOR_ICMS, '') = ''
                                                          THEN LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                          ELSE RTRIM(LTRIM(ISNULL(N.DES_VALOR_ICMS, '')))+CHAR(13)+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                      END
                             FROM #FAT_NOTAIMPRESSAO N
                                  INNER JOIN #DADOS D ON N.COD_NF = D.COD_NF
                                                         AND N.COD_SERIENOTA = D.COD_SERIENOTA
                                                         AND N.COD_FILIAL = D.COD_FILIAL
                                                         AND N.COD_REGIONAL = D.COD_REGIONAL
                             WHERE D.COD_TRIBUTACAO = @TRIBUTACAO
                                   AND D.DES_PRODUTO = @PRODUTO
                                   AND N.COD_FILIAL = @COD_FILIAL
                                   AND N.COD_REGIONAL = @COD_REGIONAL
                                   AND N.COD_NF = @COD_NF
                                   AND N.COD_SERIENOTA = @COD_SERIENOTA;
                         END;
                         ELSE
                     IF @CONTADOR = 5
                         BEGIN
                             UPDATE N
                               SET
                                   N.DES_PRODUTO_ICMS5 = CASE
                                                             WHEN LEN(ISNULL(N.DES_PRODUTO_ICMS5, '')) = 0
                                                             THEN SUBSTRING(ISNULL(D.DES_PRODUTO, ''), 1, 50)
                                                             ELSE ISNULL(N.DES_PRODUTO_ICMS5, '')
                                                         END,
                                   N.DES_VALOR_ICMS = CASE
                                                          WHEN ISNULL(N.DES_VALOR_ICMS, '') = ''
                                                          THEN LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                          ELSE RTRIM(LTRIM(ISNULL(N.DES_VALOR_ICMS, '')))+CHAR(13)+LTRIM(RTRIM(STR(D.VAL_BRUTO, 9, 2)))
                                                      END
                             FROM #FAT_NOTAIMPRESSAO N
                                  INNER JOIN #DADOS D ON N.COD_NF = D.COD_NF
                                                         AND N.COD_SERIENOTA = D.COD_SERIENOTA
                                                         AND N.COD_FILIAL = D.COD_FILIAL
                                                         AND N.COD_REGIONAL = D.COD_REGIONAL
                             WHERE D.COD_TRIBUTACAO = @TRIBUTACAO
                                   AND D.DES_PRODUTO = @PRODUTO
                                   AND N.COD_FILIAL = @COD_FILIAL
                                   AND N.COD_REGIONAL = @COD_REGIONAL
                                   AND N.COD_NF = @COD_NF
                                   AND N.COD_SERIENOTA = @COD_SERIENOTA;
                         END;
                 END;
             DELETE #DADOS
             WHERE COD_TRIBUTACAO = @TRIBUTACAO
                   AND DES_PRODUTO = @PRODUTO;
             SET @CONTADOR = @CONTADOR + 1;
         END;              
   
 --SELECT '-->' , * FROM #FAT_NOTAIMPRESSAO        
 --return       
   
 ---Lógica CFOP ---  
  
-- Carrega variáveis de configuração --  
     DECLARE @sistema INT, @versao INT;
     SELECT @sistema = des_valorcampo
     FROM FAT_CONFIGURACOESFILIALNFSE
     WHERE Cod_Regional = @COD_REGIONAL
           AND Cod_Filial = @COD_FILIAL
           AND ID_CONFIGURACOESNFSE = 29;
     SELECT @versao = des_valorcampo
     FROM FAT_CONFIGURACOESFILIALNFSE
     WHERE Cod_Regional = @COD_REGIONAL
           AND Cod_Filial = @COD_FILIAL
           AND ID_CONFIGURACOESNFSE = 30;     
---------------------------------------  
  
-- Carrega UF --  
     DECLARE @UF CHAR(2);
     SELECT @UF = E.Nom_Abvd_Estado
     FROM COR_Filial F
          INNER JOIN COR_Cidade C ON F.cod_cidade = C.cod_cidade
          INNER JOIN COR_Estado E ON C.Cod_Estado = E.Cod_Estado
     WHERE Cod_Regional = @COD_REGIONAL
           AND Cod_Filial = @COD_FILIAL;  
----------------  

     IF
(
    SELECT TP_TRIBUTACAO
    FROM #FAT_NOTAIMPRESSAO
) = 1
         BEGIN
             IF @sistema = 483
                AND @versao = 2
                 BEGIN
                     IF
(
    SELECT nom_uf
    FROM #FAT_NOTAIMPRESSAO
) <> @UF
                         BEGIN
                             UPDATE #FAT_NOTAIMPRESSAO
                               SET
                                   cod_natureza = '6933';
                         END;
                         ELSE
                         BEGIN
                             UPDATE #FAT_NOTAIMPRESSAO
                               SET
                                   cod_natureza = '5933';
                         END;
                 END;
         END;   
  
     
------------------   
   
 --*********** Insere na tabela espelho          
     INSERT INTO FAT_NOTAIMPRESSAO
(COD_FILIAL,
 COD_REGIONAL,
 COD_NF,
 COD_SERIENOTA,
 BAS_ICMS,
 BAS_ISS,
 CEP_CLIENTE,
 CEP_COBRANCA,
 CEP_FILIAL,
 CGC_CLIENTE,
 CGC_CLIENTE2,
 CGC_FILIAL,
 CGC_ORIGEM,
 CGC_SUBCLIENTENF,
 CGC_SUBCLIENTESACADO,
 CIDADE_CLIENTE,
 COD_BARRABOLETOBANC,
 COD_CGC,
 COD_CLIENTETRIBUTACAO,
 COD_EMPRESA,
 COD_FAMILIAISS,
 COD_FISCAL,
 COD_FISCAL_ISS,
 COD_FUNCCANCELAMENTO,
 COD_FUNCEMISSAO,
 COD_FUNCNFSUBS,
 COD_INTERESTADUAL,
 COD_INTERMUNICIPAL,
 COD_LINHADIGBOLETOBANC,
 COD_NATUREZA,
 COD_NFSUBSTITUTA,
 COD_RAMO,
 COD_RECIBO,
 COD_TRIBUTACAO,
 CON_CORRENTESANTANDER,
 DES_BAIRROCOBRANCA,
 DES_CANCELAMENTO,
 DES_CIDADE,
 DES_CIDADECOBRANCA,
 DES_ENDCOBRANCA,
 DES_IRRF,
 DES_JUSTFORASEQ,
 DES_MENSAGEMSUBS_ICMS,
 DES_MENSAGEMSUBS_ISS,
 DES_MSGBOLETO,
 DES_MSGNF,
 DES_NUMBOLETO,
 DES_PRODUTO_ICMS1,
 DES_PRODUTO_ICMS2,
 DES_PRODUTO_ICMS3,
 DES_PRODUTO_ICMS4,
 DES_PRODUTO_ICMS5,
 DES_PRODUTO_ISS1,
 DES_PRODUTO_ISS2,
 DES_PRODUTO_ISS3,
 DES_PRODUTO_ISS4,
 DES_PRODUTO_ISS5,
 DES_PRODUTO_ISS6,
 DES_QTD_ICMS,
 DES_QTD_ISS,
 DES_SETOREXTERNO,
 DES_SUBCLIENTE,
 DES_SUBSTRIBICMS,
 DES_TPSUBCLIENTE,
 DES_UNID_ICMS,
 DES_UNID_ISS,
 DES_VALOR_ICMS,
 DES_VALOR_ISS,
 DES_VALOREXTENSO,
 DESTACAR_ICMS,
 DESTACAR_ISS,
 DT_CANCELAMENTO,
 DT_DATA,
 DT_EMISSAO,
 DT_GRAVACAO,
 DT_NFSUBSTITUTA,
 DT_PERIFATFIM,
 DT_PERIFATINI,
 DT_VENCIMENTO,
 END_CLIENTE,
 END_FILIAL,
 FAX_FILIAL,
 FLG_AVULSA,
 FLG_CANCELAMENTO,
 FLG_CONTROLEIMPRESSA,
 FLG_DADOSSUBCLIENTENF,
 FLG_DADOSSUBCLIENTESACADO,
 FLG_ELETRONICO,
 FLG_ENVIADA,
 FLG_IMPRESSA,
 FLG_NFCOMPLEMENTAR,
 IC_CONTRIBUINTEESTADUAL,
 IC_SUBSTITUICAOTRIBUTARIA,
 IC_SUBSTITUICAOTRIBUTARIAISS,
 INSCE_CLIENTE,
 INSCE_FILIAL,
 INSCM_CLIENTE,
 INSCM_FILIAL,
 LOTEFORMULARIO,
 MSG_FISCAL,
 NOM_ABVD_FILIAL,
 NOM_CIDADE,
 NOM_CLIENTE,
 NOM_UF,
 NOM_UFCOBRANCA,
 NOMEDAGRAFICA,
 NR_CONEXAO,
 NUM_FORMULARIO1,
 NUM_FORMULARIO2,
 NUM_NF,
 PC_RETENCAIOINSS,
 REGIMEESPECIALESTADO,
 REGIMEESPECIALMUNICIPIO,
 SEQ_CLIENTE,
 SEQ_SUBCLIENTENF,
 SEQ_SUBCLIENTESACADO,
 SEQ_UTILIZA,
 TEL_FILIAL,
 TP_TRIBUTACAO,
 TOT_NOTA,
 TP_SERVICO,
 TRB_ICMS,
 TRB_ISS,
 TX_ICMS,
 TX_ISS,
 UF_FILIAL,
 UF_NOTA,
 VAL_ICMS,
 VAL_INSS,
 VAL_INSS1,
 VAL_INSS2,
 VAL_INSS4,
 VAL_INSS5,
 VAL_ISS,
 DES_PRODUTO_ISS,
 DES_PRODUTO_ICMS,
 COD_ISS,
 TX_ISS_TRIB,
 TX_ICMS_TRIB,
 ALIQ_PIS,
 ALIQ_COFINS,
 ALIQ_CSLL,
 ALIQ_IR,
 DT_VENCIMENTOAIDF,
 DES_MSGALIQUOTA,
 DES_MSGELETRONICA,
 VAL_PIS,
 VAL_COFINS,
 VAL_CSLL
 ,    
 --  <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>   
 --,Flg_IsencaoISS, Flg_IsencaoICMS , Flg_IsencaoPIS , Flg_IsencaoCOFINS , Flg_IsencaoCSLL , Flg_IsencaoIRRF )           
 Flg_IsencaoISS,
 Flg_IsencaoICMS,
 Flg_IsencaoPIS,
 Flg_IsencaoCOFINS,
 Flg_IsencaoCSLL,
 Flg_IsencaoIRRF,
 Flg_Cepom,
 Des_MsgCepom,
 UF_ORIGEM,
 id_parametroinstituicao,
 uf_destino,
 des_metodo,
 Des_CidOrigemCTe,
 Cod_IBGEOrigemCTe,
 Des_CidDestinoCTe,
 Cod_IBGEDestinoCTe,
 --<255295 - Ricardo Gomes - Plus It>
 VAL_IRRF
 --</255295 - Ricardo Gomes - Plus It>
)          
--  </alteracao>   
            SELECT COD_FILIAL,
                   COD_REGIONAL,
                   COD_NF,
                   COD_SERIENOTA,
                   BAS_ICMS,
                   BAS_ISS,
                   CEP_CLIENTE,
                   CEP_COBRANCA,
                   CEP_FILIAL,
                   CGC_CLIENTE,
                   CGC_CLIENTE2,
                   CGC_FILIAL,
                   CGC_ORIGEM,
                   CGC_SUBCLIENTENF,
                   CGC_SUBCLIENTESACADO,
                   CIDADE_CLIENTE,
                   COD_BARRABOLETOBANC,
                   COD_CGC,
                   COD_CLIENTETRIBUTACAO,
                   COD_EMPRESA,
                   COD_FAMILIAISS,
                   COD_FISCAL,
                   COD_FISCAL_ISS,
                   COD_FUNCCANCELAMENTO,
                   COD_FUNCEMISSAO,
                   COD_FUNCNFSUBS,
                   COD_INTERESTADUAL,
                   COD_INTERMUNICIPAL,
                   COD_LINHADIGBOLETOBANC,
                   COD_NATUREZA,
                   COD_NFSUBSTITUTA,
                   COD_RAMO,
                   COD_RECIBO,
                   COD_TRIBUTACAO,
                   CON_CORRENTESANTANDER,
                   DES_BAIRROCOBRANCA,
                   DES_CANCELAMENTO,
                   DES_CIDADE,
                   DES_CIDADECOBRANCA,
                   DES_ENDCOBRANCA,
                   DES_IRRF,
                   DES_JUSTFORASEQ,
                   DES_MENSAGEMSUBS_ICMS,
                   DES_MENSAGEMSUBS_ISS,
                   DES_MSGBOLETO,
                   DES_MSGNF,
                   DES_NUMBOLETO,
                   DES_PRODUTO_ICMS1,
                   DES_PRODUTO_ICMS2,
                   DES_PRODUTO_ICMS3,
                   DES_PRODUTO_ICMS4,
                   DES_PRODUTO_ICMS5,
                   DES_PRODUTO_ISS1,
                   DES_PRODUTO_ISS2,
                   DES_PRODUTO_ISS3,
                   DES_PRODUTO_ISS4,
                   DES_PRODUTO_ISS5,
                   DES_PRODUTO_ISS6,
                   DES_QTD_ICMS,
                   DES_QTD_ISS,
                   DES_SETOREXTERNO,
                   DES_SUBCLIENTE,
                   DES_SUBSTRIBICMS,
                   DES_TPSUBCLIENTE,
                   DES_UNID_ICMS,
                   DES_UNID_ISS,
                   DES_VALOR_ICMS,
                   DES_VALOR_ISS,
                   DES_VALOREXTENSO,
                   DESTACAR_ICMS,
                   DESTACAR_ISS,
                   DT_CANCELAMENTO,
                   DT_DATA,
                   DT_EMISSAO,
                   DT_GRAVACAO,
                   DT_NFSUBSTITUTA,
                   DT_PERIFATFIM,
                   DT_PERIFATINI,
                   DT_VENCIMENTO,
                   END_CLIENTE,
                   END_FILIAL,
                   FAX_FILIAL,
                   FLG_AVULSA,
                   FLG_CANCELAMENTO,
                   FLG_CONTROLEIMPRESSA,
                   FLG_DADOSSUBCLIENTENF,
                   FLG_DADOSSUBCLIENTESACADO,
                   FLG_ELETRONICO,
                   FLG_ENVIADA,
                   FLG_IMPRESSA,
                   FLG_NFCOMPLEMENTAR,
                   IC_CONTRIBUINTEESTADUAL,
                   IC_SUBSTITUICAOTRIBUTARIA,
                   IC_SUBSTITUICAOTRIBUTARIAISS,
                   INSCE_CLIENTE,
                   INSCE_FILIAL,
                   INSCM_CLIENTE,
                   INSCM_FILIAL,
                   LOTEFORMULARIO,
                   MSG_FISCAL,
                   NOM_ABVD_FILIAL,
                   NOM_CIDADE,
                   NOM_CLIENTE,
                   NOM_UF,
                   NOM_UFCOBRANCA,
                   NOMEDAGRAFICA,
                   NR_CONEXAO,
                   NUM_FORMULARIO1,
                   NUM_FORMULARIO2,
                   NUM_NF,
                   PC_RETENCAIOINSS,
                   REGIMEESPECIALESTADO,
                   REGIMEESPECIALMUNICIPIO,
                   SEQ_CLIENTE,
                   SEQ_SUBCLIENTENF,
                   SEQ_SUBCLIENTESACADO,
                   SEQ_UTILIZA,
                   TEL_FILIAL,
                   TP_TRIBUTACAO,
                   TOT_NOTA,
                   TP_SERVICO,
                   TRB_ICMS,
                   TRB_ISS,
                   TX_ICMS,
                   TX_ISS,
                   UF_FILIAL,
                   UF_NOTA,
                   VAL_ICMS,
                   VAL_INSS,
                   VAL_INSS1,
                   VAL_INSS2,
                   VAL_INSS4,
                   VAL_INSS5,
                   VAL_ISS,
                   DES_PRODUTO_ISS,
                   DES_PRODUTO_ICMS,
                   COD_ISS,
                   TX_ISS_TRIB,
                   TX_ICMS_TRIB,
                   ALIQ_PIS,
                   ALIQ_COFINS,
                   ALIQ_CSLL,
                   ALIQ_IR,
                   DT_VENCIMENTOAIDF,
                   DES_MSGALIQUOTA,
                   DES_MSGELETRONICA,
                   VAL_PIS,
                   VAL_COFINS,
                   VAL_CSLL
                   ,          
 --  <alteracao autor = "EDUARDO BOQUETTI (PLUS-IT)" data = "26/01/2015" chamado = "SDM 12699 UC01"/>   
  --, Flg_IsencaoISS, Flg_IsencaoICMS , Flg_IsencaoPIS , Flg_IsencaoCOFINS , Flg_IsencaoCSLL , Flg_IsencaoIRRF      
                   Flg_IsencaoISS,
                   Flg_IsencaoICMS,
                   Flg_IsencaoPIS,
                   Flg_IsencaoCOFINS,
                   Flg_IsencaoCSLL,
                   Flg_IsencaoIRRF,
                   Flg_Cepom,
                   Des_MsgCepom,
                   UF_ORIGEM,
                   id_parametroinstituicao,
                   uf_destino,
                   des_metodo,
                   Des_CidOrigemCTe,
                   Cod_IBGEOrigemCTe,
                   Des_CidDestinoCTe,
                   Cod_IBGEDestinoCTe,
				   --<255295 - Ricardo Gomes - Plus It>
				   VAL_IRRF
				   --</255295 - Ricardo Gomes - Plus It>
--  </alteracao>           
            FROM #FAT_NOTAIMPRESSAO;   
   
   
   
 --***** MARCA COMO REFATURAMENTO  

     DECLARE @ID_LOG_REFATURADO INT;  
      
 --OBTEM MAIOR ID  
     SELECT @ID_LOG_REFATURADO = MAX(ISNULL(M.Id_LogProcessoRefaturadoOrigem, 0))
     FROM fat_movimento M
     WHERE M.COD_REGIONAL = @COD_REGIONAL
           AND M.COD_FILIAL = @COD_FILIAL
           AND M.COD_NF = @COD_NF
           AND M.COD_SERIENOTA = @COD_SERIENOTA
           AND M.Id_LogProcessoRefaturadoOrigem > 0;  
  
 --ATUALIZA NA TABELA  
     UPDATE IMP
       SET
           IMP.Id_LogProcessoRefaturadoOrigem = @ID_LOG_REFATURADO
     FROM FAT_NOTAIMPRESSAO IMP
     WHERE IMP.COD_REGIONAL = @COD_REGIONAL
           AND IMP.COD_FILIAL = @COD_FILIAL
           AND IMP.COD_NF = @COD_NF
           AND IMP.COD_SERIENOTA = @COD_SERIENOTA;  
   
  --****************************************************************  


     SET NOCOUNT OFF;   
  
  
  
  
  
  
  
  
  