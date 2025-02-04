USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SP_DELIVERY_ENVIAR]    Script Date: 08/15/2022 17:49:28 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--declare @p21 varchar(300)
--set @p21=''
--exec SP_DELIVERY_ENVIAR '01','100',20144,'2021-06-30 00:00:00',99,'LAS TURQUESAS 455 SANTA INES','ADMIN','1',17703,'B',13594,99,0,2,'03',0,'GUILLERMO TIRADO','',0,1.5,@p21 output
--select @p21

ALTER PROC [dbo].[SP_DELIVERY_ENVIAR]
    @CODCIA CHAR(2) ,
    @NUMSER INT ,--SERIE DE COMANDA
    @NUMFAC BIGINT , --NUMERO DE COMANDA
    @FECHA DATETIME ,
    @PAGO MONEY ,
    @DIRECCION VARCHAR(150) ,
    @usuario VARCHAR(20) ,
    --@SerCom VARCHAR(3) , --SERIE DE DOCUMENTO
    --@nroCom INT , --NUMERO DE DOCUMENTO
    @SerDoc VARCHAR(3) , --SERIE DE DOCUMENTO
    @NroDoc BIGINT , --NUMERO DE FACTURA
    @Fbg CHAR(1) ,
    @codcli INT = 1 ,
    @totalfac MONEY ,
    @farjabas TINYINT , --CONSUMO O DETALLADO
    @IDREPARTIDOR INT ,
    @CODIGODOCTO CHAR(2) ,
    @DSCTO MONEY ,
    @NOMBRE_CLI VARCHAR(100) ,
    @RUC_CLI CHAR(11) ,
    @TARIFA MONEY,
    @ALL_ICBPER DECIMAL(8, 2) ,
    --@AutoNumFac INT OUT ,
    @EXITO VARCHAR(300) OUT
AS
    SET NOCOUNT ON 
    SET @EXITO = ''
    DECLARE @AutoNumFac INT ,
        @MaxNumOper INT
        
          --GRABAR CLIENTE

    IF @codcli = 0
        BEGIN
            SELECT TOP 1
                    @CODCLI = C.CLI_CODCLIE
            FROM    dbo.CLIENTES c
            WHERE   CLI_CODCIA = @CODCIA
                    AND CLI_CP = 'C'
            ORDER BY CLI_CODCLIE DESC
         
            SET @codcli = @codcli + 1
        
            INSERT  INTO dbo.CLIENTES
                    ( CLI_CODCLIE ,
                      CLI_CODCIA ,
                      CLI_CP ,
                      CLI_NOMBRE ,
                      CLI_NOMBRE_ESPOSO ,
                      CLI_CASA_DIREC ,
                      CLI_RUC_ESPOSO ,
                      CLI_ESTADO ,
                      CLI_ZONA_NEW ,
                      CLI_TRAB_DIREC         
                    )
            VALUES  ( @codcli ,
                      @CODCIA ,
                      'C' ,
                      @NOMBRE_CLI ,
                      @NOMBRE_CLI ,
                      @DIRECCION ,
                      @RUC_CLI ,
                      'A' ,
                      1 ,
                      @DIRECCION
                    )
        END
        --FIN GRABAR CLIENTE
    
    IF ( SELECT TOP 1
                ISNULL(PC.ESTADODELIVERY, '')
         FROM   dbo.PEDIDOS_CABECERA pc
         WHERE  PC.CODCIA = @CODCIA
                AND PC.NUMSER = @NUMSER
                AND PC.NUMFAC = @NUMFAC
       ) = 'E'
        BEGIN
            SET @EXITO = 'El Pedido ya fue Facturado.'
            GOTO finalizar
        END
        
        
        --OBTENIENDO LA TARIFA DEL DELIVERY
   
  
        
    --DECLARE @dscto MONEY
    --SET @dscto = 0
    DECLARE @diascre INT 
    SET @diascre = 0
    DECLARE @sec INT  --FORMA DE PAGO POR DEFECTO 1 - EFECTIVO
    DECLARE @moneda CHAR(1) 
    SET @moneda = 'S'
    SET @sec = 4
    --SET @dscto = 0
    DECLARE @MaxNumFac INT
--Variables por defecto
    DECLARE @MaxNumSec INT
    DECLARE @tipomov INT ,
        @ALL_CODART INT
    DECLARE @ALL_CODTRA INT ,
        @ALL_FLAG_EXT CHAR(1) ,
        @ALL_CODCLIE INT
    DECLARE @ALL_IMPORTE_AMORT NUMERIC(9, 2)
    DECLARE @ALL_IMPORTE NUMERIC(9, 2) ,
        @ALL_CHESER VARCHAR(3)
    DECLARE @ALL_SECUENCIA INT ,
        @ALL_IMPORTE_DOLL NUMERIC(9, 2)
    DECLARE @ALL_PRECIO NUMERIC(9, 2) ,
        @ALL_CODVEN INT ,
        @ALL_NUMGUIA NUMERIC(9, 2)
    DECLARE @ALL_FBG CHAR(1) ,
        @ALL_CP CHAR(1) ,
        @ALL_TIPDOC CHAR(2)
    DECLARE @ALL_CANTIDAD NUMERIC(9, 2) ,
        @ALL_CODBAN NUMERIC(9, 2)
    DECLARE @ALL_AUTOCON VARCHAR(40) ,
        @ALL_CHENUM NUMERIC(9, 2)
    DECLARE @ALL_CHESEC VARCHAR(1) ,
        @ALL_NUMSER CHAR(3)
    DECLARE @ALL_NETO NUMERIC(9, 2) ,
        @ALL_BRUTO NUMERIC(9, 2)
    DECLARE @ALL_IMPTO NUMERIC(9, 2)
    DECLARE @ALL_DESCTO NUMERIC(9, 2)
    DECLARE @ALL_MONEDA_CAJA CHAR(1) ,
        @ALL_MONEDA_CCM CHAR(1)
    DECLARE @ALL_MONEDA_CLI CHAR(1) ,
        @ALL_NUMDOC BIGINT
    DECLARE @ALL_LIMCRE_ANT NUMERIC(9, 2) ,
        @ALL_LIMCRE_ACT NUMERIC(9, 2)
DECLARE @ALL_SIGNO_ARM INT ,
        @ALL_CODTRA_EXT INT ,
        @ALL_SIGNO_CCM INT
    DECLARE @ALL_SIGNO_CAR INT ,
        @ALL_SIGNO_CAJA INT ,
        @ALL_NUMSER_C CHAR(3)
    DECLARE @ALL_NUMFAC_C NUMERIC(9, 2) ,
        @ALL_SERDOC INT ,
        @ALL_TIPO_CAMBIO NUMERIC(9, 2)
    DECLARE @ALL_FLETE NUMERIC(9, 2) ,
        @ALL_SUBTRA VARCHAR(20) ,
        @ALL_FACART CHAR(1)
    DECLARE @ALL_CONCEPTO VARCHAR(40) ,
        @ALL_SITUACION CHAR(1) ,
        @ALL_FLAG_SO CHAR(1)
    DECLARE @ALL_IMPG1 NUMERIC(9, 2) ,
        @ALL_IMPG2 NUMERIC(9, 2)
    DECLARE @ALL_CTAG1 CHAR(1) ,
        @ALL_CTAG2 CHAR(1) ,
        @ALL_RUC CHAR(1)
    DECLARE @ALL_CODSUNAT INT ,
        @ALL_SERIE_REC INT ,
        @ALL_NUM_RECIBO INT

    DECLARE @hora VARCHAR(12)

--1. Preguntar si el NroDoc (NumFac) existe en el allog

    IF ( SELECT COUNT(ALL_NUMFAC)
         FROM   ALLOG
         WHERE  ALL_CODCIA = @CODCIA
                AND all_fbg = @fbg
                AND all_numser = @serdoc
                AND all_numfac = @nrodoc
       ) = 0
        BEGIN
            SET @AUTONUMFAC = @NRODOC
        END
    ELSE
        BEGIN
            SELECT  @AUTONUMFAC = ISNULL(MAX(ALL_NUMFAC), 0) + 1
            FROM    ALLOG
            WHERE   ALL_CODCIA = @CodCia
                    AND all_fbg = @fbg
                    AND all_numser = @serdoc
            SET @NRODOC = @AUTONUMFAC
        END

--OBTENER CODIGO DE MESA
    DECLARE @CODMESA VARCHAR(10)
    SELECT  @CODMESA = PED_CODCLIE
    FROM    PEDIDOS
    WHERE   PED_CODCIA = @codcia
            AND PED_FECHA = @FECHA
            AND PED_NUMSER = @NUMSER
            AND PED_NUMFAC = @NUMFAC
            

    SELECT  @ALL_AUTOCON = RTRIM(LTRIM(sut_descripcion)) ,
            @ALL_SIGNO_CAR = sut_signo_car ,
            @ALL_SIGNO_CAJA = sut_signo_caja ,
            @ALL_TIPDOC = sut_tipdoc ,
            @ALL_CP = sut_cp
    FROM    sub_transa
    WHERE   sut_secuencia = @sec
            AND sut_codtra = 2401



--select *  from sub_transa where sut_secuencia=1 and sut_codtra=2401


--tabla pargen
    DECLARE @serie INT
    IF @fbg = 'F'
        BEGIN
            SELECT  @serie = par_f_serie
            FROM    pargen
            WHERE   par_codcia = @codcia
        END
    ELSE
        BEGIN
            BEGIN
                SELECT  @serie = par_b_serie
                FROM    pargen
                WHERE   par_codcia = @codcia
            END
        END
      

--tipo de cambio
    SELECT  @ALL_TIPO_CAMBIO = ISNULL(cal_tipo_cambio, 0)
    FROM    calendario
    WHERE   cal_codcia = '00'
            AND cal_Fecha = @fecha
--calculando all_bruto

    DECLARE @impigv NUMERIC(9, 4)
    DECLARE @igv AS INT
    SELECT  @igv = gen_igv
    FROM    general



    SET @impigv = ( 100.0 + @igv ) / 100.0
    SET @ALL_BRUTO = @totalfac / @impigv
    SET @ALL_IMPTO = @totalfac - @ALL_BRUTO


--select @impigv,@totalfac,@ALL_IMPTO,@igv,@ALL_BRUTO

--return





    SET @ALL_IMPORTE_AMORT = @totalfac
    SET @ALL_IMPORTE = 0
    SET @ALL_CHESER = '0'
    SET @ALL_SECUENCIA = @sec
    SET @ALL_IMPORTE_DOLL = 0
    SET @ALL_PRECIO = @totalfac
    SET @ALL_CODVEN = 0
    SET @ALL_FBG = @Fbg



--Set @ALL_CP = @allcp
--Set @ALL_TIPDOC = @alltipdoc
    SET @ALL_CANTIDAD = 0
    SET @ALL_NUMGUIA = 0
    SET @ALL_CODBAN = 0
--Set @ALL_AUTOCON = null
    SET @ALL_CHENUM = 0
    SET @ALL_CHESEC = NULL

--select 'ssd'
--return

/*
declare @dd int
exec SpFacturarComanda '01','07/10/2010','ADMIN','100',15,'20',12,'B','<r><d cp="5142" st="2" pr="15" un="UNO" sc="0"/><d cp="9063" st="1" pr="5" un="UNO" sc="1"/></r>',1,20,119,1,'S',@dd out
select @dd
*/

    SET @ALL_NUMSER = CAST(@serie AS CHAR(3))
    SET @ALL_NETO = @totalfac


--Set @ALL_BRUTO = 0

--Set @ALL_IMPTO = 0
    SET @ALL_DESCTO = @DSCTO
    SET @ALL_MONEDA_CAJA = @moneda
    SET @ALL_MONEDA_CCM = NULL
    SET @ALL_MONEDA_CLI = NULL
    SELECT  @ALL_NUMDOC = ISNULL(MAX(all_numdoc), 0)
    FROM    allog
    WHERE   all_codcia = @codcia
    SET @ALL_LIMCRE_ANT = 0
    SET @ALL_LIMCRE_ACT = 0
      


    SET @ALL_CODTRA_EXT = 2401
    SET @ALL_SIGNO_CCM = 0

    SET @TipoMov = 70
    SET @ALL_NUMSER_C = '1'
    SET @ALL_SERDOC = 0
--Set @ALL_TIPO_CAMBIO = 1
    SET @ALL_FLETE = @TARIFA

    SET @ALL_FACART = NULL
    SET @ALL_SUBTRA = @ALL_AUTOCON
    SET @ALL_SITUACION = NULL
    SET @ALL_FLAG_SO = 'A'
    SET @ALL_IMPG1 = 0
    SET @ALL_IMPG2 = 0
    SET @ALL_CTAG1 = NULL
    SET @ALL_CTAG2 = NULL


    SET @ALL_SERIE_REC = NULL
    SET @ALL_NUM_RECIBO = 0
    SET @ALL_RUC = NULL
    SET @TipoMov = 10
    SET @ALL_CODTRA = 2401
    SET @ALL_FLAG_EXT = 'N'
    SET @ALL_SIGNO_ARM = -1
    SET @ALL_CONCEPTO = 'Comanda: ' + CAST(@NUMSER AS VARCHAR(10)) + '-'
        + RTRIM(LTRIM(STR(@NUMFAC)))

    SET @ALL_CODSUNAT = @CODIGODOCTO
   -- IF @fbg = 'F'
   --       BEGIN
   --    SET @ALL_CODSUNAT = 1
   --       END
   --   ELSE
   --     BEGIN
   --       IF @fbg = 'B'
   --         BEGIN
   --           SET @ALL_CODSUNAT = 3
   --     END
   -- END



    DECLARE @NroError INT


--Obteniendo numero maximo de operacion all_numoper
    SELECT  @MaxNumOper = ISNULL(MAX(ALL_NUMOPER), 0) + 1
    FROM    [dbo].[ALLOG]
    WHERE   ALL_CODCIA = @CodCia
            AND ALL_FECHA_DIA = @Fecha

    SET @MaxNumFac = @NroDoc
    
   
    BEGIN TRY
        BEGIN TRAN
       
        
        INSERT  INTO [dbo].[ALLOG]
                ( ALL_CODCIA ,
                  ALL_FECHA_DIA ,
                  ALL_NUMOPER ,
                  ALL_CODTRA ,
                  ALL_FLAG_EXT ,
                  ALL_CODCLIE ,
                  ALL_CODART ,
                  ALL_IMPORTE_AMORT ,
                  ALL_IMPORTE ,
                  ALL_CHESER ,
                  ALL_SECUENCIA ,
                  ALL_IMPORTE_DOLL ,
                  ALL_CODUSU ,
                  ALL_PRECIO ,
                  ALL_CODVEN ,
                  ALL_FBG ,
                  ALL_CP ,
                  ALL_TIPDOC ,
                  ALL_CANTIDAD ,
                  ALL_NUMGUIA ,
                  ALL_CODBAN ,
                  ALL_AUTOCON ,
                  ALL_CHENUM ,
                  ALL_CHESEC ,
                  ALL_NUMSER ,
                  ALL_NUMFAC ,
                  ALL_FECHA_VCTO ,
                  ALL_NETO ,
                  ALL_BRUTO ,
                  ALL_GASTOS ,
                  ALL_IMPTO ,
                  ALL_DESCTO ,
                  ALL_MONEDA_CAJA ,
                  ALL_MONEDA_CCM ,
                  ALL_MONEDA_CLI ,
                  ALL_NUMDOC ,
                  ALL_LIMCRE_ANT ,
                  ALL_LIMCRE_ACT ,
                  ALL_NUM_OPER2 ,
                  ALL_SIGNO_ARM ,
                  ALL_CODTRA_EXT ,
                  ALL_SIGNO_CCM ,
                  ALL_SIGNO_CAR ,
                  ALL_SIGNO_CAJA ,
                  ALL_TIPMOV ,
                  ALL_NUMSER_C ,
                  ALL_NUMFAC_C ,
                  ALL_SERDOC ,
                  ALL_TIPO_CAMBIO ,
                  ALL_FLETE ,
                  ALL_SUBTRA ,
                  ALL_HORA ,
                  ALL_FACART ,
                  ALL_CONCEPTO ,
                  ALL_NUMOPER2 ,
                  ALL_FECHA_ANT ,
                  ALL_SITUACION ,
                  ALL_FECHA_SUNAT ,
                  ALL_FLAG_SO ,
                  ALL_IMPG1 ,
                  ALL_IMPG2 ,
                  ALL_CTAG1 ,
                  ALL_CTAG2 ,
                  ALL_CODSUNAT ,
                  ALL_FECHA_PRO ,
                  ALL_FECHA_CAN ,
                  ALL_SERIE_REC ,
                  ALL_NUM_RECIBO ,
                  ALL_RUC ,
                  ALL_MESA ,
                  ALL_DELIVERY ,
                  ALL_ICBPER
	         )
        VALUES  ( @CodCia ,
                  @Fecha ,
                  @MaxNumOper ,
                  @ALL_CODTRA ,
                  @ALL_FLAG_EXT ,
                  @codcli ,
                  0 ,
                  @ALL_IMPORTE_AMORT ,
                  @ALL_IMPORTE ,
                  @ALL_CHESER ,
                  @ALL_SECUENCIA ,
                  @ALL_IMPORTE_DOLL ,
                  @Usuario ,
                  @ALL_PRECIO ,
                  0 ,
                  @ALL_FBG ,
                  @ALL_CP ,
                  @ALL_TIPDOC ,
                  @ALL_CANTIDAD ,
                  @ALL_NUMGUIA ,
                  @ALL_CODBAN ,
                  @ALL_AUTOCON ,
                  @ALL_CHENUM ,
                  @ALL_CHESEC ,
                  @SerDoc , -- SERIE DE DOCUMENTO
                  @MaxNumFac , -- NUMERO DE DOCUMENTO
                  --@Fecha ,
                  dateadd(day,coalesce(1,0),@Fecha) ,
                  @ALL_NETO ,
                  @ALL_BRUTO ,
                  @DSCTO ,
                  @ALL_IMPTO ,
                  @ALL_DESCTO ,
                  @ALL_MONEDA_CAJA ,
                  @ALL_MONEDA_CCM ,
                  @ALL_MONEDA_CLI ,
                  @MaxNumFac ,
                  @ALL_LIMCRE_ANT ,
                  @ALL_LIMCRE_ACT ,
                  @MaxNumFac ,
                  @ALL_SIGNO_ARM ,
                  @ALL_CODTRA_EXT ,
                  @ALL_SIGNO_CCM ,
                  @ALL_SIGNO_CAR ,
                  @ALL_SIGNO_CAJA ,
                  @TipoMov ,
                  @NUMSER , --SERIE DE COMANDA
                  @NUMFAC ,--NUMERO DE COMANDA
                  @SerDoc ,
                  @ALL_TIPO_CAMBIO ,
                  @ALL_FLETE ,
                  @ALL_SUBTRA ,
                  GETDATE() ,
                  @ALL_FACART ,
                  @ALL_CONCEPTO ,
                  @MaxNumOper ,
                  @Fecha ,
                  @ALL_SITUACION ,
                  @Fecha ,
                  @ALL_FLAG_SO ,
                  @ALL_IMPG1 ,
                  @ALL_IMPG2 ,
                  @ALL_CTAG1 ,
                  @ALL_CTAG2 ,
                  @ALL_CODSUNAT ,
                  @Fecha ,
                  @Fecha ,
                  @ALL_SERIE_REC ,
                  @ALL_NUM_RECIBO ,
                  @ALL_RUC ,
                  @CodMesa ,
                  1 ,
                  @ALL_ICBPER
                )
            
        DECLARE @tbltmp TABLE
            (
              cp INT ,
              st NUMERIC(9, 2) ,
              pr NUMERIC(9, 2) ,
              un VARCHAR(20) ,
              sc INT ,
              cd NUMERIC(18, 4)
            )
        DECLARE @cp INT ,
            @st NUMERIC(9, 2) ,
            @pr NUMERIC(9, 2) ,
            @un VARCHAR(20) ,
            @sc INT ,
            @cd NUMERIC(18, 4)
        DECLARE @idoc INT ,
            @sa NUMERIC(9, 2)


        DECLARE @impto NUMERIC(9, 2) ,
            @bruto AS NUMERIC(9, 2) ,
            @total NUMERIC(9, 2)


--grabando en facart


        INSERT  INTO @tbltmp
                SELECT  P.PED_CODART ,
                        P.PED_CANTIDAD ,
                        P.PED_PRECIO ,
                        P.PED_UNIDAD ,
                        P.PED_NUMSEC ,
                        p.CANTIDAD_DELIVERY
                FROM    dbo.PEDIDOS p
                WHERE   P.PED_CODCIA = @CODCIA
                        AND P.PED_NUMSER = @NUMSER
                        AND P.PED_NUMFAC = @NUMFAC
                        AND P.PED_FECHA = @FECHA
                        AND P.PED_ESTADO = 'N'

/*
SELECT PED_CANTIDAD,CANTIDAD_DELIVERY,PED_PRECIO, * FROM dbo.PEDIDOS p WHERE p.PED_ESTADO='N'
declare @p16 varchar(300)
set @p16=''
exec SP_DELIVERY_ENVIAR '01','100',28,'20140923',109,'MZ. LL LT. 33 LOS CEDROS','ADMIN','1',1,'B',6,109,0,2,0,@p16 output
select @p16
*/

--SELECT * FROM @tbltmp
        SET @TOTAL = 0

        SELECT  @total = SUM(( st * pr ))
        FROM    @tbltmp
        WHERE   CD IS NULL
        
        
        
        SELECT  @total = ISNULL(@TOTAL, 0) + ISNULL(SUM(( CD * pr )), 0)
        FROM    @tbltmp
        WHERE   CD IS NOT NULL
   
        SET @total = @total --+ ISNULL(@TARIFA,0)   --QUITADO GTS PARA NO AGREGAR EL FLETE AL CALCULO TOTAL
        SET @impto = ROUND(@total / @impigv, 2)
        SET @bruto = @total - @impto

--CAMPO PARA DETERMINAR SI EL PRODUCTO ESTA AFECTO AL ICBPER
        DECLARE @ICBPER TINYINT ,
            @VALORICBPER DECIMAL(8, 2) ,
            @GEN_ICBPER DECIMAL(8, 2) ,
            @mICBPER DECIMAL(8, 2)
        SET @ICBPER = 0
        SET @GEN_ICBPER = 0
        SET @mICBPER = 0
    
        SELECT TOP 1
                @GEN_ICBPER = G.GEN_ICBPER
        FROM    dbo.GENERAL g
    
        DECLARE cFacArt CURSOR
        FOR
            SELECT  cp ,
                    st ,
                    pr ,
                    un ,
                    sc ,
                    cd
            FROM    @tbltmp

        OPEN cFacArt

        FETCH cFacArt INTO @cp, @st, @pr, @un, @sc, @cd

        WHILE ( @@Fetch_Status = 0 )
            BEGIN
--AFECTO AL ICBPER
                SELECT TOP 1
                        @ICBPER = COALESCE(A.ART_CALIDAD, 1)
                FROM    dbo.ARTI a
                WHERE   A.ART_CODCIA = @CODCIA
                        AND A.ART_KEY = @CP
                IF @ICBPER = 0
                    BEGIN
                        SET @mICBPER = ( @ST * @GEN_ICBPER )
                    END
	--maximo nro de secuencia
                SELECT  @Maxnumsec = ISNULL(MAX(far_numsec), 0) + 1
                FROM    facart
                WHERE   far_tipmov = @TipoMov
                        AND far_numfac = @MaxNumFac
                        AND far_fbg = @fbg

                SELECT  @hora = dbo.FnDevuelveHora(GETDATE())	--Hora actual

--obteniendo stock actual del plato
                SELECT  @sa = arm_stock
                FROM    articulo
                WHERE   arm_codcia = @codcia
                        AND arm_codart = @cp

                DECLARE @Cliente VARCHAR(30)
                DECLARE @ruc CHAR(11)

                SELECT  @Cliente = RTRIM(LTRIM(Cli_Nombre)) ,
                        @ruc = cli_ruc_esposo
                FROM    CLIEntes
                WHERE   cli_codclie = @codcli

                IF @fbg = 'B'
                    BEGIN
                        SET @ruc = ''
                    END

                DECLARE @DESCONTARSTOCK BIT ,
                    @TIPO CHAR(1)
                
                SELECT  @DESCONTARSTOCK = ISNULL(A.ART_DESCONTARSTOCK, 0) ,
                        @TIPO = RTRIM(LTRIM(A.ART_FLAG_STOCK))
                FROM    dbo.ARTI a
                WHERE   A.ART_CODCIA = @CODCIA
                        AND A.ART_KEY = @CP
                        
                    

                INSERT  INTO [dbo].[FACART]
                        ( FAR_TIPMOV ,
                          FAR_CODCIA ,
                          far_numser ,
                          far_fbg ,
                          far_numfac ,
                          FAR_NUMSEC ,
                          far_cod_sunat ,
                          FAR_CODVEN ,
                          FAR_STOCk ,
                          far_codart ,--10
                          far_cantidad ,
                          FAR_PRECIO ,
                          FAR_equiv ,
                          far_descri ,
                          far_PESO ,
                          far_signo_car ,
                          far_signo_arm ,
                          far_key_dircli ,
                          far_codclie ,
                          FAR_MONEDA ,--20
                          FAR_EX_IGV ,
                          FAR_cp ,
                          FAR_fecha_compra ,
 far_estado ,
                          FAR_estado2 ,
                          FAR_COSPRO ,
                          FAR_COSPRO_ANT ,
                          far_IMPTO ,
                          FAR_TOT_FLETE ,
                          FAR_FLETE ,--30
                          FAR_DESCTO ,
                          FAR_TOT_DESCTO ,
                          FAR_GASTOS ,
                          FAR_BRUTO ,
                          FAR_NUMDOC ,
                          far_numguia ,
                          far_serguia ,
                          FAR_pordescto1 ,
                          FAR_costeo ,
                          FAR_COSTEO_REAL , --40
                          FAR_tipo_cambio ,
                          FAR_DIAS ,
                          FAR_fecha ,
                          FAR_NUMSER_C ,
                          FAR_NUMFAC_c ,
                          FAR_NUMOPER ,
                          far_precio_neto ,
                          far_otra_cia ,
                          far_transito ,
                          far_subtra , --50
                          far_JABAS ,
                          far_UNIDADES ,
                          far_mortal ,
                          far_num_precio ,
                          FAR_ORDEN_UNIDADEs ,
                          FAR_SUBTOTAL ,
                          far_turno ,
                          far_concepto ,
                          far_codusu ,
                          FAR_HORA ,--60
                          FAR_NUM_LOTE ,
                          FAR_PEDSER ,
                          FAR_PEDFAC ,
                          far_pedsec ,
                          FAR_TIPDOC ,
                          far_fecha_pro ,
                          far_fecha_can ,
                          far_fecha_control ,
                          far_ruc ,
                          far_flag_so ,
                          FAR_NUMOPER2 ,
                          FAR_OC ,
                          FAR_DELIVERY ,
                          FAR_CANTIDAD_D ,
                          FAR_ICBPER
		                )
                VALUES  ( @TipoMov ,
                          @CodCia ,
                          @SerDoc ,
                          @fbg ,
                          @MaxNumFac ,
                          @Maxnumsec ,
                          @ALL_CODSUNAT ,
                          0 ,
                          @sa - @st ,
                          @cp ,--10
                          @st ,
                          @pr ,
                          1 ,
                          @un ,
                          0 ,
                          0 ,
                          --CASE WHEN @DESCONTARSTOCK = 1 THEN @ALL_SIGNO_ARM
                          --     ELSE 0
                          --END ,
                          0 ,
                          0 ,
                          @codcli ,
                          @moneda ,--20
                          0 ,
                          @ALL_CP ,
                          @Fecha ,
                          @ALL_FLAG_EXT ,
                          @ALL_FLAG_EXT ,
                          0 ,
                          @DSCTO ,
                          @bruto ,
                          @ALL_FLETE ,
                          0 ,--30
                          0 ,
                          @DSCTO ,
                          0 ,
                          @impto ,
                          0 ,
                          0 ,
                          0 ,
                          0 ,
                          '' ,
                          '' , --40
                          1 ,
                          1 ,
                          @Fecha ,
                          @NUMSER ,
                          @NUMFAC ,
                          @MaxNumOper ,
                          0 ,
                          '' ,
                   '' ,
                          @ALL_SUBTRA , --50
                          @farjabas ,
                          0 ,
                          0 ,
                          0 ,
                          0 ,
                          @total + coalesce(@ALL_ICBPER ,0),
                          0 ,
                          @ALL_CONCEPTO ,
                          @Usuario ,
                          @hora ,--60
                          @ALL_SECUENCIA ,
                          0 ,
                          0 ,
                          NULL ,
                          @ALL_TIPDOC ,
                          @Fecha ,
                          @Fecha ,
                          @Fecha ,
                          @ruc ,
                          'A' ,
                          @MaxNumOper ,
                          @CodMesa ,
                          1 ,
                          @cd ,
                          --CASE WHEN @ICBPER = 0 THEN @mICBPER
                          --     ELSE 0
                          --END
                          @ALL_ICBPER
                        )


             

     
                UPDATE  dbo.PEDIDOS_CABECERA
                SET     PAGO = @PAGO ,
                        DIRECCION = @DIRECCION ,
                        FACTURADO = 1 ,
                        ESTADODELIVERY = 'E' ,
                        FECHASALIDA = GETDATE() --AQUI
                        ,
                        IDREPARTIDOR = @IDREPARTIDOR
                WHERE   NUMSER = @NUMSER
                        AND NUMFAC = @NUMFAC
                        AND FECHA = @FECHA
                        AND CODCIA = @CODCIA
                
     
--ACTUALIZANDO TABLA PEDIDOS - cantidad facturada
                UPDATE  PEDIDOS
                SET     PED_FAC = PED_FAC + @st
                WHERE   PED_CODCIA = @codcia
                        AND PED_FECHA = @FECHA
                        AND PED_NUMSER = @NUMSER
                        AND PED_NUMFAC = @NUMFAC
                        AND PED_NUMSEC = @sc
           
                FETCH cFacArt INTO @cp, @st, @pr, @un, @sc, @cd
            END

        CLOSE cFacArt
        DEALLOCATE cFacArt

--GRABANDO EN CARTERA Y CARACU SEGUN SEA EL CASO
--SOLO SI EL VALOR DE ALLSIGNOCAJA ES 1 GRABA EN DICHAS TABLAS



        INSERT  INTO CARTERA
                ( CAR_CP ,
                  CAR_CODCLIE ,
                  CAR_CODCIA ,
                  CAR_SERDOC ,
                  CAR_NUMDOC ,
                  CAR_TIPDOC ,
                  CAR_IMPORTE ,
                  CAR_FECHA_INGR ,
                  CAR_FECHA_VCTO ,
                  CAR_NUM_REN ,
                  CAR_CODART ,
                  CAR_IMP_INI ,
                  CAR_SITUACION ,
                  CAR_NUMSER ,
                  CAR_NUMFAC ,
                  CAR_PRECIO ,
                  CAR_CONCEPTO ,
                  CAR_CODTRA ,
                  CAR_SIGNO_CAR ,
                  CAR_CODVEN ,
                  CAR_NUMGUIA ,
                  CAR_FBG ,
                  CAR_NOMBRE_BANCO ,
                  CAR_NUM_CHEQUE ,
                  CAR_SIGNO_CAJA ,
                  CAR_TIPMOV ,
                  CAR_NUMSER_C ,
                  CAR_NUMFAC_C ,
                  CAR_FECHA_VCTO_ORIG ,
                  CAR_COMISION ,
                  CAR_CODBAN ,
                  CAR_COBRADOR ,
                  CAR_MONEDA ,
                  CAR_SERGUIA ,
                  CAR_FECHA_SUNAT ,
                  CAR_PLACA ,
                  CAR_VOUCHER ,
                  CAR_FLAG_SO ,
                  CAR_NUMOPER ,
                  CAR_FECHA_CONTROL ,
                  CAR_FECHA_ENTREGA ,
                  CAR_FECHA_DEVO ,
                  cAR_CODUNIBKO
                        
                )

--SELECT * FROM CARTERA
        VALUES  ( @ALL_CP ,
                  @codcli ,
                  @codcia ,
                  @serie ,
                  @NroDoc ,
                  @ALL_TIPDOC ,
                  @totalfac ,
                  @fecha ,
                  DATEADD(day, @diascre, @fecha) ,
                  4 ,
                  NULL ,
                  @totalfac ,
                  '' ,
                  @serie ,
                  @MaxNumFac ,
                  0 ,
                  '' ,
                  2401 ,
                  @ALL_SIGNO_CAR ,
                  0 ,
                  '0' ,
                  @Fbg ,
                  '' ,
                  '' ,
                  @ALL_SIGNO_CAJA ,
                  10 ,
                  0 ,--CAR_NUMSER_C,
                  0 ,--CAR_NUMFAC_C,
                  DATEADD(day, @diascre, @fecha) ,--CAR_FECHA_VCTO_ORIG,
                  0 ,--CAR_COMISION,
                  0 ,--CAR_CODBAN,
                  0 ,--CAR_COBRADOR,
                  @moneda ,--CAR_MONEDA,
                  0 ,--CAR_SERGUIA,
                  @fecha ,--CAR_FECHA_SUNAT,
                  NULL ,--CAR_PLACA,
                  NULL ,--CAR_VOUCHER,
                  'A' ,--CAR_FLAG_SO,
                  @MaxNumOper ,--CAR_NUMOPER,
                  @fecha ,--CAR_FECHA_CONTROL,
                  NULL ,--CAR_FECHA_ENTREGA,
                  NULL ,--CAR_FECHA_DEVO,
                  NULL--cAR_CODUNIBKO

                        
                )


        INSERT  INTO caracu
                ( CAA_CP ,
                  CAA_CODCLIE ,
                  CAA_CODCIA ,
                  CAA_TIPDOC ,
                  CAA_FECHA ,
                  CAA_NUM_OPER ,
                  CAA_SERDOC ,
                  CAA_NUMDOC ,
                  CAA_IMPORTE ,
                  CAA_SALDO ,
                  CAA_FECHA_VCTO ,
                  CAA_CONCEPTO ,
                  CAA_SIGNO_CAR ,
                  CAA_SIGNO_CCM ,
                  CAA_ESTADO ,
                  CAA_SALDO_CAR ,
                  CAA_TOTAL ,
                  CAA_NUMSER ,
                  CAA_NUMFAC ,
                  CAA_NUMGUIA ,
                  CAA_SIGNO_CAJA ,
                  CAA_TIPMOV ,
                  CAA_FBG ,
                  CAA_HORA ,
                  CAA_CODVEN ,
                  CAA_CODUSU ,
                  CAA_NUMSER_C ,
                  CAA_NUMFAC_C ,
                  CAA_SIGNO_CAJA_REAL ,
                  CAA_NUMPLAN ,
                  CAA_NUM_CHEQUE ,
                  CAA_NOTA ,
                  CAA_SITUACION ,
                  CAA_NOMBRE ,
                  CAA_CODBAN ,
                  CAA_RECIBO ,
                  CAA_INTVEN ,
                  CAA_DIASV ,
                  CAA_DIASA ,
                  CAA_TASAV ,
                  CAA_SERGUIA ,
                  CAA_FECHA_COBRO ,
                  CAA_FLAG_SO ,
                  CAA_SERIE ,
                  CAA_TIPO_CAMBIO ,
                  CAA_CODTRA ,
                  CAA_FECHA_CONTROL
                        
                )
        VALUES  ( @ALL_CP ,
                  @codcli ,
                  @CODCIA ,
                  @ALL_TIPDOC ,
                  @fecha ,
                  @MaxNumOper ,--CAA_NUM_OPER,
                  @serie ,--CAA_SERDOC,
                  @NroDoc ,--CAA_NUMDOC, --AQUI
                  @totalfac ,--CAA_IMPORTE,
                  @totalfac ,--CAA_SALDO,
                  DATEADD(day, @diascre, @fecha) ,--CAA_FECHA_VCTO,
                  '' ,--CAA_CONCEPTO,
                  @ALL_SIGNO_CAR ,--CAA_SIGNO_CAR,
                  NULL ,--CAA_SIGNO_CCM,
                  'N' ,--CAA_ESTADO,
                  @totalfac ,--CAA_SALDO_CAR,
                  @totalfac ,--CAA_TOTAL,
                  @serie ,--CAA_NUMSER,
                  @MaxNumFac ,--CAA_NUMFAC,
                  0 ,--CAA_NUMGUIA,
                  4 ,--CAA_SIGNO_CAJA,
                  10 ,--CAA_TIPMOV,
                  @fbg ,--CAA_FBG,
                  GETDATE() ,--CAA_HORA,
                  0 ,--CAA_CODVEN,
                @Usuario ,--CAA_CODUSU,
                  0 ,--CAA_NUMSER_C,
                  0 ,--CAA_NUMFAC_C,
                  0 ,--CAA_SIGNO_CAJA_REAL,
                  0 ,--CAA_NUMPLAN,
                  '' ,--CAA_NUM_CHEQUE,
                  NULL ,--CAA_NOTA,
                  '' ,--CAA_SITUACION,
                  LEFT(@Cliente, 22) ,--CAA_NOMBRE,
                  NULL ,--CAA_CODBAN,
                  0 ,--CAA_RECIBO,
                  0 ,--CAA_INTVEN,
                  0 ,--CAA_DIASV,
                  0 ,--CAA_DIASA,
                  0 ,--CAA_TASAV,
                  0 ,--CAA_SERGUIA,
                  @fecha ,--CAA_FECHA_COBRO,
                  'A' ,--CAA_FLAG_SO,
                  0 ,--CAA_SERIE
                  @ALL_TIPO_CAMBIO ,--CAA_TIPO_CAMBIO
                  2401 ,--CAA_CODTRA
                  NULL--CAA_FECHA_CONTROL
                        
                )

                
                
    
        COMMIT TRAN
    END TRY


    BEGIN CATCH
        SET @EXITO = RTRIM(LTRIM(ERROR_MESSAGE()))
        ROLLBACK TRAN
    END CATCH


    Finalizar:
    RETURN
