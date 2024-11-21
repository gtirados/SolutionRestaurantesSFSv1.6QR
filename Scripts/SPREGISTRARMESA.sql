IF EXISTS ( SELECT TOP 1
                    S.SPECIFIC_NAME
            FROM    information_schema.routines s
            WHERE   s.ROUTINE_TYPE = 'PROCEDURE'
                    AND S.ROUTINE_NAME = 'SPREGISTRARMESA' )
    BEGIN
        DROP PROC [dbo].[SPREGISTRARMESA]
    END
GO
/*
select * from mesas
delete from mesas where mes_codmes='1D'
*/
CREATE Procedure [dbo].[SPREGISTRARMESA]
@CodCia char(2),
@CodMes varchar(10),
@Mesa varchar(40),
@CodZon INT,
@COMENSALES INT
--With Encryption
as
--Valores x Defecto
Declare @mLeft int,@mTop int
Declare @Estado char(1)
Set @Estado = 'L'
Set @mLeft = 1000
Set @mTop = 1000

IF NOT EXISTS (Select MES_CODMES 
		FROM [dbo].[MESAS] WHERE MES_CODMES = @CodMes AND MES_CODCIA = @CodCia)
	BEGIN
		INSERT INTO [dbo].[MESAS] (
			MES_CODCIA,
			MES_CODMES,
			MES_DESCRIP,
			MES_CODZON,
			MES_ESTADO,
			MES_LEFT,
			MES_TOP,
			MES_COMENSALES
		)
		VALUES (
			@CodCia,
			@CodMes,
			@Mesa,
			@CodZon,
			@Estado,
			@mLeft,
			@mTop,
			@COMENSALES
		)
	END
else
	raiserror('El Codigo de Mesa Proporcionado ya existe',16,1)

GO