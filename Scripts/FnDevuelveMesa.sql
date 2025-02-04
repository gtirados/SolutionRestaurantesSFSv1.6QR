USE [BDATOS]
GO
/****** Object:  UserDefinedFunction [dbo].[FnDevuelveMozo]    Script Date: 08/13/2024 11:11:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
select dbo.FnDevuelveMesa ('01','4A')
*/
create FUNCTION [dbo].[FnDevuelveMesa](
@CodCia char(2),
@CodMesa char(10)
)

RETURNS varchar(50) 
--With Encryption
AS  
BEGIN
Declare @Mesa varchar(40)

Select @Mesa = MES_DESCRIP
From [dbo].[MESAS]
Where MES_CODCIA = @CodCia and MES_CODMES = @CodMesa

RETURN @Mesa
	END
