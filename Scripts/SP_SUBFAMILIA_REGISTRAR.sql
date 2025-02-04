USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SP_SUBFAMILIA_REGISTRAR]    Script Date: 11/26/2024 18:45:50 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
exec SP_SUBFAMILIA_REGISTRAR '01','ENSALADAS',2
*/
ALTER PROC [dbo].[SP_SUBFAMILIA_REGISTRAR]
    (
      @CODCIA CHAR(2) ,
      @DENOMINACION VARCHAR(50) ,
      @IDFAMILIA INT,
      @DSCTO MONEY
    )
AS
    SET NOCOUNT ON 

    DECLARE @CODIGO INT


    IF NOT EXISTS ( SELECT TOP 1
                            T.TAB_TIPREG
                    FROM    dbo.TABLAS t
                    WHERE   T.TAB_TIPREG = 123
                            AND T.TAB_CODCIA = @CODCIA
                            AND T.TAB_NOMLARGO = @DENOMINACION AND t.TAB_CODART = @IDFAMILIA )
        BEGIN

            SELECT  @CODIGO = ISNULL(MAX(T.TAB_NUMTAB), 0) + 1
            FROM    dbo.TABLAS t
            WHERE   T.TAB_TIPREG = 123
                    AND T.TAB_CODCIA = @CODCIA

            INSERT  INTO dbo.TABLAS
                    ( TAB_CODCIA ,
                      TAB_TIPREG ,
                      TAB_NUMTAB ,
                      TAB_NOMLARGO ,
                      TAB_NOMCORTO ,
                      TAB_CODART ,
                      TAB_CONTABLE2 ,
                      TAB_FECHA_CONTROL ,
                      TAB_DESCUENTO
	                )
            VALUES  ( @CODCIA , -- TAB_CODCIA - char(2)
                      123 , -- TAB_TIPREG - int
                      @CODIGO , -- TAB_NUMTAB - int
                      @DENOMINACION , -- TAB_NOMLARGO - char(40)
                      LEFT(@DENOMINACION, 10) , -- TAB_NOMCORTO - char(10)
                      @IDFAMILIA , -- TAB_CODART - numeric
                      '0' , -- TAB_CONTABLE2 - varchar(50)
                      GETDATE(),  -- TAB_FECHA_CONTROL - datetime
                      @DSCTO
	                )
        END
    ELSE
        BEGIN
            RAISERROR('La Descripción Proporcionada ya existe',16,1)

        END
