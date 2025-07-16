IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'dbo' -- Validación del esquema
          AND s.ROUTINE_NAME = 'SP_SUBFAMILIA_REGISTRAR'
) -- Validación del nombre
BEGIN
    DROP PROC [dbo].[SP_SUBFAMILIA_REGISTRAR];
END;
GO
/*
exec SP_SUBFAMILIA_REGISTRAR '01','ENSALADAS',2
*/
CREATE PROC [dbo].[SP_SUBFAMILIA_REGISTRAR]
(
    @CODCIA CHAR(2),
    @DENOMINACION VARCHAR(50),
    @IDFAMILIA INT,
    @DSCTO MONEY,
    @AGREGADO BIT
)
AS
SET NOCOUNT ON;

DECLARE @CODIGO INT;


IF NOT EXISTS
(
    SELECT TOP 1
           t.TAB_TIPREG
    FROM dbo.TABLAS t
    WHERE t.TAB_TIPREG = 123
          AND t.TAB_CODCIA = @CODCIA
          AND t.TAB_NOMLARGO = @DENOMINACION
          AND t.TAB_CODART = @IDFAMILIA
)
BEGIN

    SELECT @CODIGO = ISNULL(MAX(t.TAB_NUMTAB), 0) + 1
    FROM dbo.TABLAS t
    WHERE t.TAB_TIPREG = 123
          AND t.TAB_CODCIA = @CODCIA;

    INSERT INTO dbo.TABLAS
    (
        TAB_CODCIA,
        TAB_TIPREG,
        TAB_NUMTAB,
        TAB_NOMLARGO,
        TAB_NOMCORTO,
        TAB_CODART,
        TAB_CONTABLE2,
        TAB_FECHA_CONTROL,
        tab_DESCUENTO,
        TAB_AGREGADO
    )
    VALUES
    (   @CODCIA,                 -- TAB_CODCIA - char(2)
        123,                     -- TAB_TIPREG - int
        @CODIGO,                 -- TAB_NUMTAB - int
        @DENOMINACION,           -- TAB_NOMLARGO - char(40)
        LEFT(@DENOMINACION, 10), -- TAB_NOMCORTO - char(10)
        @IDFAMILIA,              -- TAB_CODART - numeric
        '0',                     -- TAB_CONTABLE2 - varchar(50)
        GETDATE(),               -- TAB_FECHA_CONTROL - datetime
        @DSCTO, @AGREGADO);
END;
ELSE
BEGIN
    RAISERROR('La Descripción Proporcionada ya existe', 16, 1);

END;
GO