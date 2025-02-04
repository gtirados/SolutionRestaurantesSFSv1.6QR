USE [BDATOS]
GO
/****** Object:  UserDefinedFunction [dbo].[FnDevuelvePlatosConsulta]    Script Date: 10/26/2022 11:09:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
select dbo.FnDevuelvePlatosConsulta ('01','20140623','100',6)
*/
ALTER FUNCTION [dbo].[FnDevuelvePlatosConsulta]
    (
      @CodCia CHAR(2) ,
      @fecha CHAR(8) ,
      @numser CHAR(3) ,
      @numfac BIGINT 
    --  @IDFAMILIA INT
    )
RETURNS VARCHAR(4000) 
--With Encryption
AS 
    BEGIN
        DECLARE @datos VARCHAR(4000) ,
            @PLATO VARCHAR(100)
        
        DECLARE @TBLPLATOS TABLE
            (
              PLATO VARCHAR(100) ,
              INDICE INT IDENTITY
            )
        SET @datos = ''
        INSERT  INTO @TBLPLATOS
                ( PLATO
                )
                SELECT  A.ART_NOMBRE
                FROM    dbo.PEDIDOS p
                        INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                 AND P.PED_CODART = A.ART_KEY
                WHERE   P.PED_CODCIA = @CodCia
                        AND P.PED_NUMSER = @NUMSER
                        AND P.PED_NUMFAC = @numfac
                        AND P.PED_ESTADO = 'N'
                        AND CONVERT(VARCHAR(8), P.PED_FECHA, 112) = @fecha --AND p.PED_FAMILIA2 = @IDFAMILIA
                        
        DECLARE @MIN INT ,
            @MAX INT
        SELECT  @MIN = MIN(T.INDICE)
        FROM    @TBLPLATOS t
        SELECT  @MAX = MAX(T.INDICE)
        FROM    @TBLPLATOS t
                        
        WHILE @MIN <= @MAX 
            BEGIN
                SELECT  @PLATO = T.PLATO
                FROM    @TBLPLATOS t
                WHERE   T.INDICE = @MIN
							
                IF @MIN = @MAX 
                    BEGIN
                        SET @datos = @datos + ' ' + RTRIM(LTRIM(@PLATO))
                    END
                ELSE 
                    BEGIN
                        SET @datos = @datos + ' ' + RTRIM(LTRIM(@PLATO)) + ','
                    END
							
                SET @MIN = @MIN + 1
            END

        RETURN @datos
    END
