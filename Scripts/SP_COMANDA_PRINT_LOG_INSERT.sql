USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SP_COMANDA_PRINT_LOG_INSERT]    Script Date: 11/06/2024 22:06:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
exec SP_COMANDA_PRINT_LOG_INSERT '01','2024-06-17 00:00:00',163296,'SUPERVISOR','CARLOS HUIMAN'
*/
ALTER PROC [dbo].[SP_COMANDA_PRINT_LOG_INSERT]
    (
      @CODCIA CHAR(2) ,
      @FECHA DATE ,
      @NUMFAC INT ,
      @USUARIO VARCHAR(20) ,
      @MOZO VARCHAR(40)
    )
AS
    SET NOCOUNT ON 
    DECLARE @TOTAL MONEY ,
        @IDLOG BIGINT ,
        @IDMESA VARCHAR(10)

    SELECT TOP 1
            @IDLOG = CL.IDLOG + 1
     FROM    dbo.CUENTAPRINT_LOG cl
    WHERE   CL.CODCIA = @CODCIA
    ORDER BY CL.IDLOG DESC

    IF @IDLOG IS NULL
        BEGIN
            SET @IDLOG = 1
        END
        
    SELECT TOP 1
            @IDMESA = P.CODMESA
    FROM    dbo.PEDIDOS_CABECERA p
    WHERE   P.FECHA = @FECHA
            AND P.NUMFAC = @NUMFAC
            AND P.CODCIA = @CODCIA
            

    SELECT  @TOTAL = SUM(P.PED_SUBTOTAL)
    FROM    dbo.PEDIDOS p
    WHERE   P.PED_FECHA = @FECHA
            AND P.PED_NUMFAC = @NUMFAC
            AND P.PED_CODCIA = @CODCIA
            AND P.PED_ESTADO = 'N'

    INSERT  INTO dbo.CUENTAPRINT_LOG
            ( CODCIA ,
              IDLOG ,
              PEDNUMFAC ,
              IDMESA ,
              MOZO ,
              FECHA ,
              USUARIO ,
              TOTAL
            )
    VALUES  ( @CODCIA , -- CODCIA - char(2)
              @IDLOG , -- IDLOG - bigint
              @NUMFAC , -- PEDNUMFAC - int
              @IDMESA , -- MESA - varchar(20)
              @MOZO , -- MOZO - varchar(20)
              --GETDATE() , -- FECHA - datetime
              @FECHA,
              @USUARIO , -- USUARIO - varchar(20)
              @TOTAL -- TOTAL - money
            )

