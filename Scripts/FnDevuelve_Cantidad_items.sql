alter FUNCTION [dbo].[FnDevuelve_Cantidad_items]
    (
      @CODCIA CHAR(2) ,
      @FECHA DATETIME ,
      @NUMFAC BIGINT ,
      @NUMSER VARCHAR(3) ,
      @IDFAMILIA INT
    )
RETURNS INT
--With Encryption
AS
    BEGIN
        DECLARE @MIN INT ,
            @MAX INT
    
        
        DECLARE @CANT INT
        SET @CANT  =0
    
				SELECT   @CANT= COUNT(p.PED_CODART)
                              FROM      dbo.PEDIDOS p
                                        INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY  AND A.ART_FLAG_STOCK <> 'C'
                              WHERE     P.PED_CODCIA = @CODCIA
                                        AND P.PED_ESTADO = 'N'
                                        AND P.PED_FECHA = @FECHA
                                        AND P.PED_NUMSER = @NUMSER
                                        AND P.PED_NUMFAC = @NUMFAC
                                        AND P.PED_CANATEN <> P.PED_CANTIDAD
                                        --AND p.PED_FAMILIA2 = @IDFAMILIA
           
                                        IF EXISTS (
                                        SELECT  TOP 1 a.ART_KEY
                              FROM      dbo.PEDIDOS p
                                        INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY  AND A.ART_FLAG_STOCK = 'C'
                              WHERE     P.PED_CODCIA = @CODCIA
                                        AND P.PED_ESTADO = 'N'
                                        AND P.PED_FECHA = @FECHA
                                        AND P.PED_NUMSER = @NUMSER
                                        AND P.PED_NUMFAC = @NUMFAC
                                        AND P.PED_CANATEN <> P.PED_CANTIDAD
                                        --AND p.PED_FAMILIA2 = @IDFAMILIA
                                        )
                                        BEGIN
											SELECT @CANT=@CANT + COUNT(PA_CODCIA) + 
											(SELECT COUNT(PX.PED_CODCIA) FROM PEDIDOS px
INNER JOIN ARTI ax ON px.PED_CODCIA = ax.ART_CODCIA AND px.PED_CODART = ax.ART_KEY AND ax.ART_FLAG_STOCK='C'
WHERE PED_CODCIA=@CODCIA AND PED_NUMFAC=@NUMFAC AND PED_NUMSER = @NUMSER AND PED_ESTADO='N')
											FROM PAQUETES 
											INNER JOIN ARTI ON PAQUETES.PA_CODART = ARTI.ART_KEY AND PAQUETES.PA_CODCIA = ARTI.ART_CODCIA
											WHERE PA_CODCIA = @CODCIA AND PA_CODPA IN(SELECT a.ART_KEY
                              FROM      dbo.PEDIDOS p
                                        INNER JOIN dbo.ARTI a ON P.PED_CODCIA = A.ART_CODCIA
                                                              AND P.PED_CODART = A.ART_KEY  AND A.ART_FLAG_STOCK = 'C'
                              WHERE     P.PED_CODCIA = @CODCIA
                                        AND P.PED_ESTADO = 'N'
                                        AND P.PED_FECHA = @FECHA
                                        AND P.PED_NUMSER = @NUMSER
                                        AND P.PED_NUMFAC = @NUMFAC
                          AND P.PED_CANATEN <> P.PED_CANTIDAD
                                        --AND p.PED_FAMILIA2 = @IDFAMILIA
                                        )
           AND ARTI.ART_FAMILIA = @IDFAMILIA
                                        END
                                     
                                     
                                      
    
        
        RETURN @CANT
    END




