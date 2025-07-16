IF EXISTS ( SELECT TOP 1
                    S.SPECIFIC_NAME
            FROM    information_schema.routines s
            WHERE   s.ROUTINE_TYPE = 'PROCEDURE'	-- Validación del tipo
            		AND ROUTINE_SCHEMA = 'dbo'		-- Validación del esquema
                    AND S.ROUTINE_NAME = 'SP_SUBFAMILIA_SEARCH' )		-- Validación del nombre
    BEGIN
        DROP PROC [dbo].[SP_SUBFAMILIA_SEARCH]
    END
GO
/*
SP_SUBFAMILIA_SEARCH '01'
*/
CREATE PROC [dbo].[SP_SUBFAMILIA_SEARCH]
    (
      @CODCIA CHAR(2) ,
      @IDFAMILIA INT = -1 ,
      @SEARCH VARCHAR(50) = NULL
    )
	WITH ENCRYPTION
AS
    SET NOCOUNT ON 
    
    SELECT  T.TAB_NUMTAB AS 'IDE' ,
            T.TAB_NOMLARGO AS 'NOM' ,
            sf.TAB_NOMLARGO AS 'FAMILIA' ,
            T.TAB_CODART AS 'IDEFAMILIA',
            T.TAB_DESCUENTO AS 'DESCUENTO'
			,COALESCE(t.TAB_AGREGADO,0) AS 'AGREGADO'
    FROM    dbo.TABLAS t
            INNER JOIN dbo.TABLAS SF ON t.TAB_CODCIA = sf.TAB_CODCIA
                                        AND t.TAB_CODART = sf.TAB_NUMTAB
                                        AND SF.TAB_TIPREG = 122
    WHERE   t.TAB_TIPREG = 123
            AND T.TAB_CODCIA = @CODCIA
            AND ( T.TAB_NOMLARGO LIKE '%' + @SEARCH + '%'
                  OR @SEARCH IS NULL
                )
            AND ISNULL(T.TAB_CODART, -1) = CASE WHEN ISNULL(@IDFAMILIA, -1) = -1
                                                THEN ISNULL(T.TAB_CODART, -1)
                                                ELSE @IDFAMILIA
                                           END
GO