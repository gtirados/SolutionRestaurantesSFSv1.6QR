IF EXISTS
(
    SELECT TOP 1
        s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SPLISTARMESAS'
)
BEGIN
    DROP PROC [dbo].[SPLISTARMESAS];
END;
GO
/*
select * from mesas
SpListarMesas '01'
*/
CREATE PROCEDURE [dbo].[SPLISTARMESAS] @CodCia AS CHAR(2)
--With Encryption
AS
SET NOCOUNT ON;
SELECT 
m.MES_CODMES AS 'CodMesa',
       LTRIM(RTRIM(m.MES_DESCRIP)) AS 'Mesa',
       m.MES_CODZON AS 'CodZona',
       dbo.FnDevuelveZona(@CodCia, m.MES_CODZON) AS 'Zona',
	   m.MES_COMENSALES AS 'comensales',
       m.MES_ESTADO AS 'Estado'
FROM [dbo].[MESAS] m
WHERE MES_CODCIA = @CodCia
ORDER BY Zona,
         MES_DESCRIP;

GO