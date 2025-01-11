IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_FORMASPAGO_VALIDA'
)
BEGIN
    DROP PROC [dbo].[USP_FORMASPAGO_VALIDA];
END;
GO
/*
USP_FORMASPAGO_VALIDA '01','b','3',56
USP_FORMASPAGO_VALIDA '01','b','3',57
*/
CREATE PROCEDURE [dbo].[USP_FORMASPAGO_VALIDA]
    @CODCIA CHAR(2),
   @TIPODOCTO CHAR(1),
   @SERIE VARCHAR(3),
   @NUMERO BIGINT
WITH ENCRYPTION
AS
SET NOCOUNT ON;
DECLARE @MASDEUNO BIT
SET @MASDEUNO = 0

IF (SELECT COUNT(cp.NUMERO) FROM COMPROBANTE_FORMAPAGO cp WHERE CODCIA = @CODCIA AND TIPODOCTO = @TIPODOCTO AND SERIE = @SERIE AND NUMERO = @NUMERO)> 1
BEGIN
	SET @MASDEUNO = 1	
END

SELECT @MASDEUNO AS 'masdeuno'
GO