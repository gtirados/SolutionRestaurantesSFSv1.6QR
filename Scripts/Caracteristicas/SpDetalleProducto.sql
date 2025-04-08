/*
exec SpDetalleProducto '01',157,1
*/
IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SpDetalleProducto'
)
BEGIN
    DROP PROC [dbo].[SpDetalleProducto];
END;
go
/*
SpDetalleProducto '01',74292,3
SpDetalleProducto '01',74185,2
*/
CREATE proc [dbo].[SpDetalleProducto]
@codcia char(2),
@codart bigint,
@TotPa int
as
set nocount on


select a.art_nombre as 'Prod',p.pa_prom as 'Prom',@TotPa * p.pa_prom as 'Total',a.ART_KEY as 'codigo'
from arti a
inner join paquetes p on
a.art_codcia = p.pa_codcia and a.art_key = p.pa_codart
where p.pa_codcia = @codcia and p.pa_codpa = @codart

