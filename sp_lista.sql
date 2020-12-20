Create proc sp_lista.sql

as

select * from tbRetorno where Dt_retorno > Getate()

GO


