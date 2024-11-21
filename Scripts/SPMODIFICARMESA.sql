IF EXISTS
(
    SELECT TOP 1
        s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SPMODIFICARMESA'
)
BEGIN
    DROP PROC [dbo].[SPMODIFICARMESA];
END;
GO
/*
SpModificarMesa '01','A8'
*/
CREATE PROCEDURE [dbo].[SPMODIFICARMESA]
    @CodCia CHAR(2),
    @CodMes VARCHAR(10),
    @Mesa VARCHAR(40),
    @CodZon INT,
    @COMENSALES INT
--With Encryption
AS
IF NOT EXISTS
(
    SELECT MES_CODZON
    FROM [dbo].[MESAS]
    WHERE MES_CODZON = @CodZon
          AND MES_CODCIA = @CodCia
          AND MES_DESCRIP = @Mesa
          AND MES_CODMES <> @CodMes
)
BEGIN
    UPDATE [dbo].[MESAS]
    SET MES_DESCRIP = @Mesa,
        MES_CODZON = @CodZon,
        MES_COMENSALES = @COMENSALES
    WHERE MES_CODCIA = @CodCia
          AND MES_CODMES = @CodMes;

END;
ELSE
    RAISERROR('La Denominación de la Mesa Proporcionado ya existe', 16, 1);

GO