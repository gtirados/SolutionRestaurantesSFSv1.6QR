USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SP_DELIVERY_DATOS_CLIENTE]    Script Date: 03/11/2024 11:31:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[SP_DELIVERY_DATOS_CLIENTE]
    @CodCia CHAR(2) ,
    @NUMSER CHAR(3) ,
    @NUMFAC BIGINT ,
    @Fecha DATETIME 

--With encryption
AS --Obtener la ultima comanda de acuerdo al valor maximo de ped_numfac

    SET NOCOUNT ON 
      
    SELECT  ISNULL(PC.DIRECCION, '') AS 'DIRECCION' ,
            ISNULL(C.CLI_NOMBRE, '') AS 'CLIENTE' ,
            ISNULL(PC.IDCLIENTE, -1) AS 'IDECLIENTE' ,
            ISNULL(PC.PAGO, 0) AS 'PAGO' ,
            ISNULL(PC.VUELTO, 0) AS 'VUELTO' ,
            ISNULL(pc.DESCUENTO, 0) AS 'DESCUENTO',
            ISNULL(c.CLI_RUC_ESPOSA, '') as 'DNI',
            ISNULL(c.CLI_RUC_ESPOSO, '') as 'RUC'
    FROM    dbo.PEDIDOS_CABECERA pc
            LEFT JOIN dbo.CLIENTES c ON PC.IDCLIENTE = C.CLI_CODCLIE
                                        AND PC.CODCIA = C.CLI_CODCIA
                                        AND C.CLI_ESTADO = 'A'
                                        AND c.CLI_CP = 'C'
    WHERE   PC.CODCIA = @CodCia
            AND PC.NUMFAC = @NUMFAC
            AND PC.NUMSER = @NUMSER
