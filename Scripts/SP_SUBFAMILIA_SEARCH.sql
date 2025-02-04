USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SP_SUBFAMILIA_SEARCH]    Script Date: 11/26/2024 20:41:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
SP_SUBFAMILIA_SEARCH '01'
*/
ALTER PROC [dbo].[SP_SUBFAMILIA_SEARCH]
    (
      @CODCIA CHAR(2) ,
      @IDFAMILIA INT = -1 ,
      @SEARCH VARCHAR(50) = NULL
    )
AS
    SET NOCOUNT ON 
    
    SELECT  T.TAB_NUMTAB AS 'IDE' ,
            T.TAB_NOMLARGO AS 'NOM' ,
            sf.TAB_NOMLARGO AS 'FAMILIA' ,
            T.TAB_CODART AS 'IDEFAMILIA',
            T.TAB_DESCUENTO AS 'DESCUENTO'
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
