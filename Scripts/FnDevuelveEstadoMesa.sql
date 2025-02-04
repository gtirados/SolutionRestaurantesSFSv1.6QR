USE [BDATOS]
GO
/****** Object:  UserDefinedFunction [dbo].[FnDevuelveMesa]    Script Date: 11/06/2024 09:12:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
select dbo.FnDevuelveEstadoMesa ('01','5')
*/
create FUNCTION [dbo].[FnDevuelveEstadoMesa](
@CodCia char(2),
@CodMesa char(10)
)

RETURNS varchar(50) 
--With Encryption
AS  
BEGIN
Declare @Mesa varchar(40)

Select @Mesa = MES_ESTADO
From [dbo].[MESAS]
Where MES_CODCIA = @CodCia and MES_CODMES = @CodMesa

RETURN @Mesa
	END
