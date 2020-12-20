Create proc sp_lista.sql

as

select * from tbRetorno where Dt_retorno > Getate()

GO

update tbRetorno set Dv_enviado = 'S' where Dv_enviado  = 'N'

GO


