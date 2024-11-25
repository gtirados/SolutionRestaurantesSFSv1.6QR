
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
/*
exec SpCargarFormasPago 2401
*/
ALTER PROC [dbo].[SPCARGARFORMASPAGO]
    @CodTran INT ,
    @DELIVERY BIT = 0
AS
    SET NOCOUNT ON 
	
    IF @DELIVERY = 1
        BEGIN
            SELECT  ST.SUT_SECUENCIA AS 'CODIGO' ,
                    RTRIM(LTRIM(ST.SUT_DESCRIPCION)) AS 'FORMAPAGO'
            FROM    dbo.SUB_TRANSA st
            WHERE   ST.sut_codtra = @CodTran
                    AND ST.SUT_SIGNO_CAJA = 1
        END
    ELSE
        BEGIN
            SELECT  ST.SUT_SECUENCIA AS 'CODIGO' ,
                    RTRIM(LTRIM(ST.SUT_DESCRIPCION)) AS 'FORMAPAGO' ,
                    SUT_TIPO AS 'TIPO',
                    SUT_SIGNO_CAR AS 'CRE'
            FROM    dbo.SUB_TRANSA st
            WHERE   ST.sut_codtra = @CodTran
        END
    

