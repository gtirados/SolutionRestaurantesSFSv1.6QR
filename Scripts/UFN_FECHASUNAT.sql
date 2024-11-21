
CREATE FUNCTION UFN_FECHASUNAT(@FECHA DATE)
RETURNS VARCHAR(10)
AS
BEGIN

DECLARE @RESULTADO VARCHAR(10)
SET @RESULTADO = CAST(YEAR(@FECHA) AS VARCHAR(4)) + '-' + RIGHT('00'
                                                              + CAST(MONTH(@FECHA) AS VARCHAR(2)),
                                                              2) + '-'
            + RIGHT('00' + CAST(DAY(@FECHA) AS VARCHAR(2)), 2)
            
            RETURN @RESULTADO
END