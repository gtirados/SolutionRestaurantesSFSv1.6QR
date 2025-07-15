/*
select dbo.UFN_CONCATENARFORMASPAGO(all_codcia, all_numser, all_numfac, all_fbg), ALL_FBG, ALL_NUMSER, ALL_NUMFAC,* from allog where all_tipmov = 10
select*from COMPROBANTE_FORMAPAGO
select * from SUB_TRANSA where SUT_CODTRA = 2401
*/

create function UFN_CONCATENARFORMASPAGO(@CODCIA CHAR(2),@NUMSER VARCHAR(3), @NUMFAC BIGINT, @FBG CHAR(1))
RETURNS VARCHAR(500)
AS
BEGIN
	DECLARE @resultado varchar(max) = '';
	
	--select @resultado = @resultado +
	--case when @resultado = '' then '' else ' - ' end + RTRIM(LTRIM(fp.SUT_DESCRIPCION))
	
	select @resultado = @resultado +  
	RTRIM(LTRIM(fp.SUT_DESCRIPCION)) + ' => ' + cast(cfp.monto as varchaR(20)) + char(13) + char(10)
	
	from allog a inner join comprobante_Formapago cfp on a.all_codcia= cfp.codcia
	and a.all_numser = cfp.serie and a.all_numfac = cfp.numero and a.all_fbg = cfp.tipodocto
	inner join Sub_Transa fp on cfp.idformapago = fp.sut_secuencia AND fp.SUT_CODTRA='2401'
	where a.all_fbg = @fbg and a.all_numser = @numser and a.all_numfac = @numfac and a.all_codcia= @codcia
	
	return @resultado
END


--select fp.*
--	from allog a inner join comprobante_Formapago cfp on a.all_codcia= cfp.codcia
--	and a.all_numser = cfp.serie and a.all_numfac = cfp.numero and a.all_fbg = cfp.tipodocto
--	inner join Sub_Transa fp on cfp.idformapago = fp.sut_secuencia AND fp.SUT_CODTRA='2401'
--	where a.all_fbg = 'B' and a.all_numser = '3' and a.all_numfac = 1 and a.all_codcia= '01'