IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'dbo' -- Validación del esquema
          AND s.ROUTINE_NAME = 'SPMODIFICARCOMANDA1'
) -- Validación del nombre
BEGIN
    DROP PROC [dbo].[SPMODIFICARCOMANDA1];
END;
GO
/*

*/
CREATE PROCEDURE [dbo].[SPMODIFICARCOMANDA1]
    @CodCia CHAR(2),
    @Usuario VARCHAR(10),
    @CodMesa VARCHAR(10),
    @cp INT,
    @cant NUMERIC(18, 4), --JULIO 11-01-2011  
    @pre NUMERIC(18, 4),  --JULIO 11-01-2011  
    @imp NUMERIC(18, 4),  --JULIO 11-01-2011  
    @d VARCHAR(50),
    @Mozo INT,
    @NumSer CHAR(3),
    @NumFac INT,
    @NUMSEC INT OUT,
    @Fecha DATETIME,
    @CodFam INT,
    @CLIENTE VARCHAR(120),
    @COMENSALES INT,
    @PADRE INT = NULL
--With Encryption  
AS --Valores por Defecto  
SET NOCOUNT ON;
DECLARE @PedEstado CHAR(1);

DECLARE @ValIgv NUMERIC(18, 4),
        @Moneda CHAR(1),
        @FBG CHAR(1); --JULIO VALIGV NUMERIC(18,4) 11-01-2011  
DECLARE @TipMov INT,
        @Transp INT,
        @Condi INT,
        @Dias INT;
DECLARE @DirCli INT;
SET @DirCli = 0;
SET @Transp = 0;
SET @Condi = 1;
SET @Dias = 0;
SET @FBG = 'B';
SET @Moneda = 'S';
SET @PedEstado = 'N';
SET @TipMov = 201;
SET @ValIgv = 1.18;

--===================  
DECLARE @Mesa VARCHAR(50);
DECLARE @igv NUMERIC(18, 4); --JULIO 11-01-2011  
DECLARE @PreEqui INT,
        @Unidad VARCHAR(12),
        @Hora VARCHAR(12);

--OBTENGO LA MESA  
SELECT TOP 1
       @CodMesa = pc.CODMESA
FROM dbo.PEDIDOS_CABECERA pc
WHERE pc.FECHA = @Fecha
      AND pc.CODCIA = @CodCia
      AND pc.NUMSER = @NumSer
      AND pc.NUMFAC = @NumFac
      AND pc.FACTURADO = 0;

/*  
 SELECT  @NUMSEC = ISNULL(MAX(ped_numsec) , 0) + 1  
 FROM    pedidos  
 WHERE   ped_codcia = @codcia  
 AND ped_numser = @numser  
 AND ped_numfac = @numfac  
    */

--02-01-2013  
SELECT TOP 1
       @NUMSEC = p.PED_NUMSEC + 1
FROM dbo.PEDIDOS p
WHERE p.PED_NUMFAC = @NumFac
      AND p.PED_CODCIA = @CodCia
      AND p.PED_NUMSER = @NumSer
ORDER BY p.PED_NUMSEC DESC;

--SELECT @NUMSEC AS 'NUMSEC'  
--Select * from pedidos  
--Obteniendo el Nombre de la MEsa  
SELECT @Mesa = RTRIM(LTRIM(MES_DESCRIP))
FROM [dbo].[MESAS]
WHERE MES_CODCIA = @CodCia
      AND MES_CODMES = @CodMesa;

--Obteniendo el precio equivalente  
SELECT @PreEqui = PRE_EQUIV,
       @Unidad = PRE_UNIDAD
FROM [dbo].[PRECIOS]
WHERE PRE_CODCIA = @CodCia
      AND PRE_CODART = @cp;
--Obteniendo la Hora actual  
SET @Hora = dbo.FnDevuelveHora(GETDATE());
--Disgregando el IGV  
SET @igv = @imp - ROUND((@imp / @ValIgv), 2);
INSERT INTO [dbo].[PEDIDOS]
(
    PED_CODCIA,
    PED_FECHA,
    PED_NUMSER,
    PED_NUMFAC,
    PED_NUMSEC,
    PED_CANTIDAD,
    PED_PRECIO,
    PED_CODUSU,
    PED_IGV,
    PED_BRUTO,
    PED_ESTADO,
    PED_CODART,
    PED_UNIDAD,
    PED_EQUIV,
    PED_CODCLIE,
    PED_OFERTA,
    PED_TIPMOV,
    PED_HORA,
    PED_MONEDA,
    PED_NOMCLIE,
    PED_SUBTOTAL,
    PED_FBG,
    PED_TRANSP,
    PED_CONDI,
    PED_DIAS,
    PED_CODVEN,
    PED_DIRCLI,
    PED_APROBADO,
    PED_FAMILIA,
    PED_FAMILIA2,
    PED_CLIENTE,
    PED_COMENSALES,
    PED_FECHAREG
	,PED_PADRE
)
VALUES
(   @CodCia,
    --  convert(varchar(10),getdate(),103),  
    @Fecha, @NumSer, @NumFac, @NUMSEC, @cant, @pre, @Usuario, @igv, @imp - @igv, @PedEstado, @cp, @Unidad, @PreEqui,
    @CodMesa, @d, @TipMov, @Hora, @Moneda, @Mesa, @imp, @FBG, @Transp, @Condi, @Dias, @Mozo, @DirCli, '0', @CodFam,
    @CodFam, @CLIENTE, @COMENSALES, GETDATE(),@PADRE);


UPDATE dbo.PEDIDOS
SET PED_CLIENTE = @CLIENTE,
    PED_COMENSALES = @COMENSALES
WHERE PED_NUMSER = @NumSer
      AND PED_NUMFAC = @NumFac
      AND PED_FECHA = @Fecha
      AND PED_CODCIA = @CodCia;
GO