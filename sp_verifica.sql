Create procedure sp_verifica
as



select * from tbErrosProcessos order by Dt_erros

GO


update tbTransmite 

Go

select * from tbRetorno

Go

update tbRetorno set Dt_retorno = Getdate() where Dv_enviado = ''

GO

