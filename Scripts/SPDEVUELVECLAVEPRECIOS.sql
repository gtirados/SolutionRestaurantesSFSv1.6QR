USE [BDATOS]
GO
/****** Object:  StoredProcedure [dbo].[SPDEVUELVECLAVECAJA]    Script Date: 08/03/2024 16:17:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
SPDEVUELVECLAVEPRECIOS @USUARIO='ADMIN',@CLAVE='fatima'
*/
CREATE PROCEDURE [dbo].[SPDEVUELVECLAVEPRECIOS]
    @USUARIO VARCHAR(10) ,
    @CLAVE VARCHAR(10) ,
    @MSN VARCHAR(200) OUT ,
    @PASA BIT OUT
AS 
    SET nocount ON
    
    SET @MSN = ''
            --1. VALIDANDO Q EL USUARIO EXISTE
    IF EXISTS ( SELECT  u.USU_NOMBRE
                FROM    usuarios u
                WHERE   usu_key = @USUARIO ) 
        BEGIN
            IF EXISTS ( SELECT  u.USU_NOMBRE
                        FROM    usuarios u
                        WHERE   usu_key = @USUARIO
                                AND USU_CLAVE = @CLAVE
                                AND USU_CAMBIAPRECIOS = 'A' ) 
                BEGIN
                    IF EXISTS ( SELECT  u.USU_NOMBRE
                                FROM    usuarios u
                                WHERE   usu_key = @USUARIO
                                        AND USU_CLAVE = @CLAVE
                                        AND USU_CAMBIAPRECIOS = 'A' ) 
                        BEGIN
                            SET @PASA = 1            	
                        END
                    ELSE 
                        BEGIN
                            SET @PASA = 0
                            SET @MSN='La clave proporcionada es incorrecta.'
                        END
                    
                END
            ELSE 
                BEGIN
                    SET @MSN = 'No Tiene los permisos Necesarios para efectuar la operación.'
                    SET @PASA = 0
                END
        END
    ELSE 
        BEGIN
            SET @MSN = 'Usuario no registrado en el Sistema.'
            SET @PASA = 0
        END
