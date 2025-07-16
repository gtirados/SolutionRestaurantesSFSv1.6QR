IF EXISTS ( SELECT TOP 1
                    S.SPECIFIC_NAME
            FROM    information_schema.routines s
            WHERE   s.ROUTINE_TYPE = 'PROCEDURE'	-- Validación del tipo
            		AND ROUTINE_SCHEMA = 'dbo'		-- Validación del esquema
                    AND S.ROUTINE_NAME = 'SP_SUBFAMILIA_MODIFICAR' )		-- Validación del nombre
    BEGIN
        DROP PROC [dbo].[SP_SUBFAMILIA_MODIFICAR]
    END
GO
/*

*/
CREATE PROC [dbo].[SP_SUBFAMILIA_MODIFICAR]
(
    @CODCIA CHAR(2),
    @DENOMINACION VARCHAR(50),
    @IDFAMILIA INT,
    @DSCTO MONEY,
    @AGREGADO BIT,
    @CODIGO INT
)
AS
SET NOCOUNT ON;

UPDATE dbo.TABLAS
SET TAB_NOMLARGO = @DENOMINACION,
    TAB_CODART = @IDFAMILIA,
    tab_DESCUENTO = @DSCTO,
    TAB_AGREGADO = @AGREGADO
WHERE TAB_TIPREG = 123
      AND TAB_CODCIA = @CODCIA
      AND TAB_NUMTAB = @CODIGO;

GO