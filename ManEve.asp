<%@ Language ="Vbscript"%>
<% Response.Expires = 0 %>
<!-- #INCLUDE FILE="../../INCLUDE/REDIR2.TXT" -->
<!--
################################
#Página para cadastrar/alterar evento contábeis
#Alterado por: Carlos Souza
#Data: 5/09/2000
#Alterações:	Adicionar um campo para que o usuário possa digitar o número do evento
#				contábil
################################
-->
<html>

<head>
<title>Evento</title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
</head>
<!-- #INCLUDE FILE="../../INCLUDE/Help.TXT" -->
<!-- #INCLUDE FILE="../../INCLUDE/Format.TXT" -->

<body background="../../Figuras/FUNDO.GIF">
<form name="ManEve" method="post">

<%

n_Emp		    = request.querystring("n_Emp")
i_Evento	    = request.querystring("i_Evento")
t_Evento	    = request.querystring("t_Evento")
Evento		    = request.querystring("Evento")
Dscr		    = request.querystring("Dscr")
b_Calc_IRRF		= request.querystring("b_Calc_IRRF")
b_Calc_ISS		= request.querystring("b_Calc_ISS")
b_Calc_INSS		= request.querystring("b_Calc_INSS")
b_Calc_CSLL		= request.querystring("b_Calc_CSLL")
b_Calc_COFINS	= request.querystring("b_Calc_COFINS")
b_Calc_PIS		= request.querystring("b_Calc_PIS")
b_Trib_Evento	= request.querystring("b_Trib_Evento")
i_Trib_IR_PF	= request.querystring("i_Trib_IR_PF")
i_Trib_IR_PJ	= request.querystring("i_Trib_IR_PJ")
n_Nat_Rend_PF   = request.querystring("n_Nat_Rend_PF")
n_Nat_Rend_PJ   = request.querystring("n_Nat_Rend_PJ")
b_Serv          = Request.QueryString ("b_Serv")
t_Doc_Fiscal    = Request.QueryString ("t_Doc_Fiscal")
n_Nat_Serv      = Request.QueryString ("n_Nat_Serv")
x_Funcao        = Request.QueryString ("x_Funcao")
Funcao			= request.querystring("Funcao")

session("ChnEmp") = n_Emp

Set DbObj = Server.CreateObject("ADODB.Connection")
%>

<!-- #INCLUDE FILE="../../INCLUDE/Conexao.ASP" -->

<%

	SQL = "exec spr_lst_empresa '" & n_Emp & "',1"
    SET oRS = DbObj.Execute(SQL)

	if not oRs.eof then Emp = Trim(oRs("Emp"))


	SQL = "exec spr_lst_par_pagnet " & n_Emp & ",1,0,1"
    SET oRS = DbObj.Execute(SQL)

	b_Exib_Trib_Evento = 0
	if not oRs.eof then	b_Exib_Trib_Evento = oRs("b_Exib")


	SQL = "exec spr_lst_par_pagnet " & n_Emp & ",1024,0,1"
    SET oRS = DbObj.Execute(SQL)
	b_Exib_Cod_Trib = 0
	if not oRs.eof then	b_Exib_Cod_Trib = oRs("b_Exib")


	SQL = "exec spr_lst_par_pagnet " & n_Emp & ",64,0,1"
    SET oRS = DbObj.Execute(SQL)

	b_Exib_Entrada_i_Evento = 0
	if not oRs.eof then	b_Exib_Entrada_i_Evento = oRs("b_Exib")


    b_Exib_Novos_Impostos = 0
    SQL = "exec spr_lst_par_pagnet " & session("nEmp") & ",262144,0,1"
    SET oRS = DbObj.Execute(SQL)
    if not oRs.eof then	b_Exib_Novos_Impostos = oRs("b_Exib")
    session("bExibNovosImpostos") = b_Exib_Novos_Impostos


		'Inclusão
		session("FEvento") = 1
		session("ChiEvento") = ""


	if request.form("optbCalcIRRF") <> "" then Funcao = ""
	x_Funcao = 0
    t_Inf_Rend_PF = 0
	
	if request.form("cbotEvento") = "" and Funcao = "" AND i_Evento <> "" then

		SQL = "exec spr_lst_evento " & n_Emp & ", '" & i_Evento & "' ,1"
		SET oRs = DbObj.Execute(SQL)

		'Inclusão
		'session("FEvento") = 1
		'session("ChiEvento") = ""
      
		if not oRs.eof then
		
			session("bCalcIRRFANT") = ""
			session("bCalcISSFANT") = ""
			session("bCalcINSSFANT") = ""
			session("bCalcCSLLFANT") = ""
			session("bCalcCOFINSFANT") = ""
			session("bCalcPISFANT") = ""
				
			for i = 0 to oRs.fields.count - 1
				if oRS.fields(i).name = "x_Funcao" then					
					if not isnull(oRs("x_Funcao")) then x_Funcao = Trim(oRs("x_Funcao"))
					exit for
				end if	
			next		

			t_Evento		= Trim(oRs("t_Evento"))
			Evento			= Trim(oRs("Evento"))
			if not isnull(oRs("b_Calc_IRRF")) then b_Calc_IRRF = oRs("b_Calc_IRRF") else b_Calc_IRRF = 1
			if not isnull(oRs("b_Calc_ISS")) then b_Calc_ISS = oRs("b_Calc_ISS") else b_Calc_ISS = 1
			if not isnull(oRs("b_Calc_INSS")) then b_Calc_INSS = oRs("b_Calc_INSS") else b_Calc_INSS = 1
			b_Trib_Evento	= oRs("b_Trib_Evento")
			if not isnull(oRs("i_Trib_IR_PF")) then i_Trib_IR_PF = oRs("i_Trib_IR_PF") else i_Trib_IR_PF = 0
			if not isnull(oRs("i_Trib_IR_PJ")) then i_Trib_IR_PJ = oRs("i_Trib_IR_PJ") else i_Trib_IR_PJ = 0
			if not isnull(oRs("b_Calc_CSLL")) then b_Calc_CSLL = oRs("b_Calc_CSLL") else b_Calc_CSLL = 1
			if not isnull(oRs("b_Calc_COFINS")) then b_Calc_COFINS = oRs("b_Calc_COFINS") else b_Calc_COFINS = 1
			if not isnull(oRs("b_Calc_PIS")) then b_Calc_PIS = oRs("b_Calc_PIS") else b_Calc_PIS = 1
			if not isnull(oRs("b_Serv")) then b_Serv = Trim(oRs("b_Serv"))
			
			if not isnull(oRs("t_Doc_Fiscal")) then t_Doc_Fiscal = Trim(oRs("t_Doc_Fiscal"))
			if not isnull(oRs("n_Nat_Serv")) then n_Nat_Serv = Trim(oRs("n_Nat_Serv"))
			if not isnull(oRs("n_Nat_Rend_PF")) then 
				n_Nat_Rend_PF = Trim(oRs("n_Nat_Rend_PF"))
			else
				n_Nat_Rend_PF = "-1"
			end if
			
			if not isnull(oRs("n_Nat_Rend_PJ")) then 
				n_Nat_Rend_PJ = Trim(oRs("n_Nat_Rend_PJ"))
			else
				n_Nat_Rend_PJ = "-1"
			end if			

			if not isnull(oRs("t_Inf_Rend_PF")) then t_Inf_Rend_PF = Trim(oRs("t_Inf_Rend_PF"))
			'Alteração
			session("FEvento") = 2
			session("ChiEvento") = i_Evento

			'Guarda a informacao anterior a respeito do calculo de impostos
			if b_Calc_IRRF = 5 OR b_Calc_IRRF = 6 then session("bCalcIRRFANT") = b_Calc_IRRF
			if b_Calc_ISS = 5 OR b_Calc_ISS = 6 then session("bCalcISSFANT") = b_Calc_ISS
			if b_Calc_INSS = 5 OR b_Calc_INSS = 6 then session("bCalcINSSFANT") = b_Calc_INSS
			if b_Calc_CSLL = 5 OR b_Calc_CSLL = 6 then session("bCalcCSLLFANT") = b_Calc_CSLL
			if b_Calc_COFINS = 5 OR b_Calc_COFINS = 6 then session("bCalcCOFINSFANT") = b_Calc_COFINS
			if b_Calc_PIS = 5 OR b_Calc_PIS = 6 then session("bCalcPISFANT") = b_Calc_PIS
						
		end if

		if i_Evento <> "" then
			SQL = "exec spr_lst_help_evento " & n_Emp & ", " & i_Evento
			SET oRs = DbObj.Execute(SQL)
			if not oRs.eof then	Dscr = Trim(oRs("Dscr"))
		end if
	else
			t_Evento		= request.form("cbotEvento")
			Evento			= request.form("txtEvento")
			Dscr			= request.form("txtDscr")

			b_Trib_Evento	= 0
			i_Trib_IR_PF	= 0
			i_Trib_IR_PJ	= 0

			if b_Exib_Trib_Evento = 1 then
			
				
				IF Request.Form("chkbCalcIRRFPF") <> "" then b_Calc_IRRF = 2
				IF Request.Form("chkbCalcIRRFPJ") <> "" then b_Calc_IRRF = 3
				IF Request.Form("chkbCalcIRRFPF") <> "" AND Request.Form("chkbCalcIRRFPJ") <> "" THEN b_Calc_IRRF = 4				
				IF Request.Form ("chkbCalcIRRFInter") <> "" then b_Calc_IRRF = 5
				IF Request.Form ("chkbCalcIRRFRetInter") <> "" then b_Calc_IRRF = 6
				
				
				IF Request.Form("chkbCalcISSPF") <> "" then b_Calc_ISS = 2
				IF Request.Form("chkbCalcISSPJ") <> "" then b_Calc_ISS = 3
				IF Request.Form("chkbCalcISSPF") <> "" AND Request.Form("chkbCalcISSPJ") <> "" THEN b_Calc_ISS = 4
				IF Request.Form ("chkbCalcISSInter") <> "" then b_Calc_ISS = 5
				IF Request.Form ("chkbCalcISSRetInter") <> "" then b_Calc_ISS = 6
				
				IF Request.Form("chkbCalcINSSPF") <> "" then b_Calc_INSS = 2
				IF Request.Form("chkbCalcINSSPJ") <> "" then b_Calc_INSS = 3
				IF Request.Form("chkbCalcINSSPF") <> "" AND Request.Form("chkbCalcINSSPJ") <> "" THEN b_Calc_INSS = 4
				IF Request.Form ("chkbCalcINSSInter") <> "" then b_Calc_INSS = 5
				IF Request.Form ("chkbCalcINSSRetInter") <> "" then b_Calc_INSS = 6
									
			end if

	end if


	if request.form("optbTribEvento") <> "" then
		b_Trib_Evento	= request.form("optbTribEvento")
		i_Trib_IR_PF	= request.form("cboiTribIRPF")
		i_Trib_IR_PJ	= request.form("cboiTribIRPJ")
	end if	



	if b_Exib_Trib_Evento = 0 then
		if TRIM(b_Calc_IRRF) = "" then b_Calc_IRRF = 1
		if TRIM(b_Calc_ISS)  = "" then b_Calc_ISS = 1
		if TRIM(b_Calc_INSS) = "" then b_Calc_INSS = 1
		if TRIM(b_Calc_CSLL) = "" then b_Calc_CSLL = 1
		if TRIM(b_Calc_COFINS) = "" then b_Calc_COFINS = 1
		if TRIM(b_Calc_PIS) = "" then b_Calc_PIS = 1
	else
		if TRIM(b_Calc_IRRF) = "" then b_Calc_IRRF = 0
		if TRIM(b_Calc_ISS)  = "" then b_Calc_ISS = 0
		if TRIM(b_Calc_INSS) = "" then b_Calc_INSS = 0	
		if TRIM(b_Calc_CSLL) = "" then b_Calc_CSLL = 0
		if TRIM(b_Calc_COFINS) = "" then b_Calc_COFINS = 0
		if TRIM(b_Calc_PIS) = "" then b_Calc_PIS = 0

	end if
	
	if i_Trib_IR_PF		= "" then i_Trib_IR_PF = 0
	if i_Trib_IR_PJ		= "" then i_Trib_IR_PJ = 0
	if b_Serv = "" then b_Serv = 0

    if b_Trib_Evento = "" then b_Trib_Evento = "0"

	session("bExibTribEvento") = b_Exib_Trib_Evento
	session("bExibCodTrib") = b_Exib_Cod_Trib
	session("bExibEntradaiEvento") = b_Exib_Entrada_i_Evento
	session("bCalcIRRF") = b_Calc_IRRF
	
	

dim FN(10)

xFuncao = Clng(x_Funcao)
	
For i = 1 to 10
	FN(i) = ""		
next
		
For i = 10 to 1 step-1 
	if xFuncao >= 2 ^ (i - 1) then	
			xFuncao = xFuncao - 2 ^(i - 1)
			FN(i) = "checked"				
	end if
Next	

%>


<table border="0" width="100%" style="border-bottom: 1px solid <% =session("CorLine") %>" cellspacing="0">
  <tr>
    <td width="100%"><font face="<% =session("Fonte") %>" size="4" color="<% =session("CorST") %>"><em><strong>
	<p style="margin-top: 1px"></strong>Evento</em></font>
	</td>
  </tr>
</table>

<table border="0" width="100%" cellspacing="0" style="margin-top: 5px">
  <tr>
    <td width="18%"><font face="<% =session("Fonte") %>" size="2">
	<p style="margin-top: 1px"></strong>Empresa:</font>
	</td>
    <td width="82%"><font face="<% =session("Fonte") %>" size="2" color="<% =session("CorST") %>">
	<p style="margin-top: 1px"></strong><% =n_Emp %> - <% =Emp %></font>
  </tr>

  <tr>
    <td width="18%"><font face="<% =session("Fonte") %>" size="2">
	<p style="margin-top: 1px"></strong>Código:</font>
	</td>
    <td width="82%"><font face="<% =session("Fonte") %>" size="2" color="<% =session("CorST") %>">
	
	<% 
	if b_Exib_Entrada_i_Evento = 1 then
		
		if session("FEvento") = 1 then 
			'Insere uma caixa de texto para o usuario digitar o codigo do evento.
%>			
			<p style='margin-top: 1px'></strong><input type='text' size='10' name='User_Code_Evento'></font>			
<%      else 
%>
			<p style='margin-top: 1px'></strong><% =i_Evento %></font>
<%      end if 
		
	else 

		if session("FEvento") = 1 then 
%>
			<p style='margin-top: 1px'></strong>NOVO</font>
<%		else 
%>
			<p style='margin-top: 1px'></strong><% =i_Evento %></font>
<% 
		end if 

	end if 
%>
  </tr>
</table>

<table border="0" width="100%" cellspacing="0" style="margin-top: 10px">  
  <tr>
    <td width="18%"><font face="<% =session("Fonte") %>" size="2">Tipo de Evento:</font></td>
    <td width="82%"><font face="<% =session("Fonte") %>" size="2">
    <% 'if session("FEvento") = "2" then Aux = "disabled" else Aux = "" 
    %>
	<select name="cbotEvento" <%=Aux%> size="1" style="font-family: monospace; font-size: 12pt">

<%

	SQL = "exec spr_lst_tipo_evento 0,1"
	SET oRs = DbObj.Execute(SQL)

	while not oRs.eof 

		if t_Evento = "" then t_Evento = Trim(oRs("t_Evento"))					
		
		if t_Evento = Trim(oRs("t_Evento")) then
%>		
			<option value='<% =oRs("t_Evento") %>' selected><% =oRs("t_Evento") %><%=Ret_Esp(3,oRs("t_Evento")) %><% =trim(oRs("Dscr")) %>
<%
		else
%>		
			<option value='<% =oRs("t_Evento") %>'><% =oRs("t_Evento") %><%=Ret_Esp(3,oRs("t_Evento")) %><% =trim(oRs("Dscr")) %>
<%
		end if
		
		oRs.movenext
	Wend 

%>


    </select></font></td>
  </tr>
  <tr>
    <td width="18%"><font face="<% =session("Fonte") %>" size="2">Evento:</font></td>
    <td width="82%"><font face="<% =session("Fonte") %>" size="2"></font>
	<input type="text" name="txtEvento" size="58" value="<%=Evento%>" maxlength="60" style="text-transform: uppercase"><font face="<% =session("Fonte") %>" size="2"></font></td>
  </tr>
</table>

<table border='0' width='100%' cellspacing='0' style='margin-top: 15px'>
<tr>
<td width='100%'><p style='margin-top: 1px; margin-bottom: 1px'>
<font face='<% =session("Fonte") %>' size='1'>CALCULAR IMPOSTOS</font></td></tr>
</table>

<table border='0' width='100%' style='margin-top: 1px; border: 1px solid rgb(0,0,0)' cellspacing='0'>
<tr>
<td width='100%'>

	<table border='0' width='100%'>
	<tr>				
	<% 
	if b_Exib_Trib_Evento = 0 then %>
		<td width='15%'><font face='<% =session("Fonte") %>' size='2'>IRRF?</font></td>	
		<td width='85%'>
		<input type='radio' value='1' name='optbCalcIRRF' onclick="ExibNat()"><font face='<% =session("Fonte") %>' size='2'>Sim</font>
		<input type='radio' value='0' name='optbCalcIRRF' onclick="ExibNat()"><font face='<% =session("Fonte") %>' size='2'>Não</font>
		<input type='radio' value='5' name='optbCalcIRRF' onclick="ExibNat()"><font face='<% =session("Fonte") %>' size='2'>Valor Informado</font>
		<input type='radio' value='6' name='optbCalcIRRF' onclick="ExibNat()"><font face='<% =session("Fonte") %>' size='2'>Valor Retido na Interface</font>
		</td>
		</tr>
		<tr>
		<td width='15%'><font face='<% =session("Fonte") %>' size='2'>ISS?</font></td>
		<td width='85%' align='left'>
		<input type='radio' value='1' name='optbCalcISS'><font face='<% =session("Fonte") %>' size='2'>Sim</font>
		<input type='radio' value='0' name='optbCalcISS'><font face='<% =session("Fonte") %>' size='2'>Não</font>
		<input type='radio' value='5' name='optbCalcISS'><font face='<% =session("Fonte") %>' size='2'>Valor Informado</font>
		</td>		
		</tr>
		<tr>
		<td width='15%'><font face='<% =session("Fonte") %>' size='2'>INSS?</font></td>
		<td width='85%'>
		<input type='radio' value='1' name='optbCalcINSS'><font face='<% =session("Fonte") %>' size='2'>Sim</font>
		<input type='radio' value='0' name='optbCalcINSS'><font face='<% =session("Fonte") %>' size='2'>Não</font>
		<input type='radio' value='5' name='optbCalcINSS'><font face='<% =session("Fonte") %>' size='2'>Valor Informado</font>
		</td>
		</tr>
		<%
		if b_Exib_Novos_Impostos <> 0 then
		%>
			<tr>
			<td width='15%'><font face='<% =session("Fonte") %>' size='2'>CSLL?</font></td>
			<td width='85%'>
			<%
			if b_Calc_CSLL = "1" then Aux = "checked" else Aux = "" %>
			<input type='radio' value='1' name='optbCalcCSLL' <%=Aux%>><font face='<% =session("Fonte") %>' size='2'>Sim</font>
			<%
			if b_Calc_CSLL = "0" then Aux = "checked" else Aux = "" %>
			<input type='radio' value='0' name='optbCalcCSLL' <%=Aux%>><font face='<% =session("Fonte") %>' size='2'>Não</font>
			</td>
			</tr>		
			<tr>
			<td width='15%'><font face='<% =session("Fonte") %>' size='2'>COFINS?</font></td>
			<td width='85%'>
			<%
			if b_Calc_COFINS = "1" then Aux = "checked" else Aux = "" %>
			<input type='radio' value='1' name='optbCalcCOFINS' <%=Aux%>><font face='<% =session("Fonte") %>' size='2'>Sim</font>
			<%
			if b_Calc_COFINS = "0" then Aux = "checked" else Aux = "" %>
			<input type='radio' value='0' name='optbCalcCOFINS' <%=Aux%>><font face='<% =session("Fonte") %>' size='2'>Não</font>
			</td>
			</tr>		
			<tr>
			<td width='15%'><font face='<% =session("Fonte") %>' size='2'>PIS?</font></td>
			<td width='85%'>
			<%
			if b_Calc_PIS = "1" then Aux = "checked" else Aux = "" %>
			<input type='radio' value='1' name='optbCalcPIS' <%=Aux%>><font face='<% =session("Fonte") %>' size='2'>Sim</font>
			<%
			if b_Calc_PIS = "0" then Aux = "checked" else Aux = "" %>
			<input type='radio' value='0' name='optbCalcPIS' <%=Aux%>><font face='<% =session("Fonte") %>' size='2'>Não</font>
			</td>
			</tr>		
		<%
		end if
		%>
	<% 
	else ' b_Exib_Trib_Evento <> 0  
	%>
	
		<td width='15%'><font face='<% =session("Fonte") %>' size='2'>IRRF?</font></td>		
		<td width='85%'><INPUT type='checkbox' value='2' name='chkbCalcIRRFPF' ONCLICK='ATUTELA("IRRF2")'><font face='<% =session("Fonte") %>' size='2'>PF</font>
		<INPUT type='checkbox' value='3' name='chkbCalcIRRFPJ' ONCLICK='ATUTELA("IRRF3")'><font face='<% =session("Fonte") %>' size='2'>PJ</font>
        <INPUT type='checkbox' value='6' name='chkbCalcIRRFRetInter' ONCLICK='ATUTELA("IRRF6")'><font face='<% =session("Fonte") %>' size='2'>Valor Retido na Interface</font>
		<INPUT type='checkbox' value='5' name='chkbCalcIRRFInter' ONCLICK='ATUTELA("IRRF7")'><font face='<% =session("Fonte") %>' size='2'>Valor Informado</font>		
		</td>
		</tr>
		<tr>
		<td width='15%'><font face='<% =session("Fonte") %>' size='2'>ISS?</font></td>
		<td width='85%'><INPUT type='checkbox' value='2' name='chkbCalcISSPF' ONCLICK='ATUTELA("ISS2")'><font face='<% =session("Fonte") %>' size='2'>PF</font>
		<INPUT type='checkbox' value='3' name='chkbCalcISSPJ' ONCLICK='ATUTELA("ISS3")'><font face='<% =session("Fonte") %>' size='2'>PJ</font>
		<INPUT type='checkbox' value='6' name='chkbCalcISSRetInter' ONCLICK='ATUTELA("ISS6")'><font face='<% =session("Fonte") %>' size='2'>Valor Retido na Interface</font>
		<INPUT type='checkbox' value='5' name='chkbCalcISSInter' ONCLICK='ATUTELA("ISS5")'><font face='<% =session("Fonte") %>' size='2'>Valor Informado</font>
		</td>		
		</tr>
		<tr>		
		<td width='15%'><font face='<% =session("Fonte") %>' size='2'>INSS?</font></td>
		<td width='15%'><INPUT type='checkbox' value='2' name='chkbCalcINSSPF' ONCLICK='ATUTELA("INSS2")'><font face='<% =session("Fonte") %>' size='2'>PF</font>
		<INPUT type='checkbox' value='3' name='chkbCalcINSSPJ' ONCLICK='ATUTELA("INSS3")'><font face='<% =session("Fonte") %>' size='2'>PJ</font>
		<INPUT type='checkbox' value='6' name='chkbCalcINSSRetInter' ONCLICK='ATUTELA("INSS6")'><font face='<% =session("Fonte") %>' size='2'>Valor Retido na Interface</font>
		<INPUT type='checkbox' value='5' name='chkbCalcINSSInter' ONCLICK='ATUTELA("INSS5")'><font face='<% =session("Fonte") %>' size='2'>Valor Informado</font>
		</td>
		</tr>
		<tr>
		<% if b_Calc_CSLL = "3" then Aux = "checked" else Aux = "" %>		
		<td width='15%'><font face='<% =session("Fonte") %>' size='2'>CSLL?</font></td>
		<td width='85%'><INPUT type='checkbox' value='3' name='chkbCalcCSLLPJ' <%=Aux%>><font face='<% =session("Fonte") %>' size='2'>PJ</font>
		</td>		
		</tr>		
		<tr>
		<% if b_Calc_COFINS = "3" then Aux = "checked" else Aux = "" %>
		<td width='15%'><font face='<% =session("Fonte") %>' size='2'>COFINS?</font></td>
		<td width='85%'><INPUT type='checkbox' value='3' name='chkbCalcCOFINSPJ' <%=Aux%>><font face='<% =session("Fonte") %>' size='2'>PJ</font>
		</td>		
		</tr>		
		<tr>
		<% if b_Calc_PIS = "3" then Aux = "checked" else Aux = "" %>
		<td width='15%'><font face='<% =session("Fonte") %>' size='2'>PIS?</font></td>
		<td width='85%'><INPUT type='checkbox' value='3' name='chkbCalcPISPJ' <%=Aux%>><font face='<% =session("Fonte") %>' size='2'>PJ</font>
		</td>		
		</tr>		
	<% 
	end if 
	
	if session("bExibIsenDoencaEvento") = "1" AND ( b_Calc_IRRF = "1" or b_Calc_IRRF = "2" or b_Calc_IRRF = "4") then AUX = "" else AUX = "display:none" 
%>
	<tr id="DivDoencaGrave" style="<%=AUX%>">
	<td width='15%'><font face='<% =session("Fonte") %>' size='2'>Isenção doença grave:</font></td>		
	<td width='85%'>
	<% if (x_Funcao and 8) = 8 then AUX = "checked" else AUX = "" %>
	<input type='radio' value='8' name='optbDoencaGrave' <%=AUX%>><font face='<% =session("Fonte") %>' size='2'>Sim
	<% if (x_Funcao and 8) = 8 then AUX = "" else AUX = "checked" %>
	<input type='radio' value='0' name='optbDoencaGrave' <%=AUX%>><font face='<% =session("Fonte") %>' size='2'>Não
	</font>
	</td>
	</tr>	
	
		<tr id="DivNatRendPF">
			<td><font face='<% =session("Fonte") %>' size='2'>Nat Rend default PF:</font></td>
			<td>
				<select name='cbonNatRendPF' size='1' style='font-family: monospace; font-size: 8pt'>
					<option value='-1'>DEFAULT DO CADASTRO</option>
				<% 
					SQL = "exec spr_lst_nat_rend @Parametro = '1', @nAcao = 4"
					SET oRs = DbObj.Execute(SQL)	
					oRs.Filter = "n_Nat_Rend <> 0"
					While not oRs.eof
						
						if Trim(oRs("n_Nat_Rend")) = Trim(n_Nat_Rend_PF) then 
							AUX="selected" 
						else 
							AUX= ""
						end if
				%>
							<option value="<% =oRs("n_Nat_Rend") %>" <%=AUX%>><% = oRs("n_Nat_Rend") & " - " & oRs("Dscr") %></option>
				<% 
						oRs.movenext
					Wend 
				%>	
			    </select></td>
			 </td>
		</tr>
		
		<% if n_Nat_Rend_PF = "561" then AUX = "" else AUX = "display:none" %>		
		
		<tr id="DivInfRendPF" style="<%=AUX%>">
			<td><font face='<% =session("Fonte") %>' size='2'>Informe de Rendimento:</font></td>
			<td>
				<select name='cbotInfRendPF' size='1' style='font-family: monospace; font-size: 8pt'>
					<% if t_Inf_Rend_PF = "0" then AUX = "SELECTED" else AUX = "" %>
					<option value='0' <%=AUX%>>DE ACORDO COM A NATUREZA DE RENDIMENTO</option>
					<% if t_Inf_Rend_PF = "1" then AUX = "SELECTED" else AUX = "" %>
					<option value='1' <%=AUX%>>PESSOA FÍSICA</option>
					<% if t_Inf_Rend_PF = "2" then AUX = "SELECTED" else AUX = "" %>
					<option value='2' <%=AUX%>>FINANCEIRO</option>
			    </select></td>
			 </td>
		</tr>
		
		<tr id="DivNatRendPJ">
			<td><font face='<% =session("Fonte") %>' size='2'>Nat Rend default PJ:</font></td>
			<td>
				<select name='cbonNatRendPJ' size='1' style='font-family: monospace; font-size: 8pt'>
					<option value='-1'>DEFAULT DO CADASTRO</option>
				<% 
					SQL = "exec spr_lst_nat_rend @Parametro = '2', @nAcao = 4"
					SET oRs = DbObj.Execute(SQL)	
					oRs.Filter = "n_Nat_Rend <> 0"
					While not oRs.eof
					
						if Trim(oRs("n_Nat_Rend")) = Trim(n_Nat_Rend_PJ) then 
							AUX="selected" 
						else 
							AUX= ""
						end if
				%>
							<option value="<% =oRs("n_Nat_Rend") %>" <%=AUX%>><% = oRs("n_Nat_Rend") & " - " & oRs("Dscr") %></option>
				<% 
						oRs.movenext
					Wend 
				%>	
			    </select>
			 </td>
		</tr>						
	
    </tr>

    </table>
</td>			
</tr>
<% 
if b_Exib_Cod_Trib = 1 then 
%>
	<tr>
	<td width='100%'>

		<table border='0' width='100%' cellspacing='0' style='margin-top: 5px'>
		<tr>
		<td width='40%'><font face='<% =session("Fonte") %>' size='2'>Utilizar Código de Tributação do IR: </font></td>
		<td width='60%'><select name='cbobTribEvento' size='1' style='font-family: monospace; font-size: 8pt'>		
		<% if b_Trib_Evento = "0" then Aux = "selected" else Aux = "" %>
		<option value='0' <%=Aux%>>Cadastro</option>
		<% if b_Trib_Evento = "1" then Aux = "selected" else Aux = "" %>
		<option value='1' <%=Aux%>>Evento</option>
		</tr>
		<tr>
		<td width='40%'><font face='<% =session("Fonte") %>' size='2'>Código Tributação PF: </td>
		<td width='60%'>
		<select name='cboiTribIRPF' size='1' style='font-family: monospace; font-size: 8pt'>
		<option value=''>
	<% 
		SQL = "exec spr_lst_tab_ir_pf 0,1"
		SET oRs = DbObj.Execute(SQL)	

		While not oRs.eof
		
			if Trim(oRs("i_Trib_IR_PF")) = Trim(i_Trib_IR_PF) then 
	%>
				<option value=<% =oRs("i_Trib_IR_PF") %> selected><% =oRs("Dscr") %></option>
	<% 
			else 
	%>
				<option value=<% =oRs("i_Trib_IR_PF") %>><% =oRs("Dscr") %></option>
	<% 
			end if 
			oRs.movenext
		Wend 
	%>	
	    </select></td>
		</tr>
		<tr>					
		<td width='40%'><font face='<% =session("Fonte") %>' size='2'>Código Tributação PJ: </td>
		<td width='60%'>
		<select name='cboiTribIRPJ' size='1' style='font-family: monospace; font-size: 8pt'>
	    <option value=''>
	<% 
		SQL = "exec spr_lst_tab_ir_pj 0,1"
		Set oRs = DbObj.Execute(SQL)

		While not oRs.eof
			if Trim(oRs("i_Trib_IR_PJ")) = Trim(i_Trib_IR_PJ) then 
	%>
				<option value=<% =oRs("i_Trib_IR_PJ") %> selected><% =oRs("Dscr") %></option>
	<% 
			else 
	%>
				<option value=<% =oRs("i_Trib_IR_PJ") %>><% =oRs("Dscr") %></option>
	<% 
			end if 
			oRs.movenext
		Wend 
	%> 
		</select></td>
		</tr>
		</table>
	</td>
	</tr>  

<%
end if 'b_Exib_Cod_Trib = 1

Response.Write "x_Funcao = " & x_Funcao

%>
	<tr>
	<td width='100%'>
		<table border='0' width='100%' cellspacing='0' style='margin-top: 5px'>
		<tr>
		<td width="18%"><font face="<% =session("Fonte") %>" size="2">Evento de Pensão Alimentícia:</font></td>
		<td width="82%"><font face="<% =session("Fonte") %>" size="2"></font>
		<% if (x_Funcao and 32) = 32 then AUX = "checked" else AUX = "" %>
		<input type='radio' value='32' name='optxFuncao' <%=AUX%> ><font face='<% =session("Fonte") %>' size='2'>Sim
		<% if (x_Funcao and 32) = 32 then AUX = "" else AUX = "checked" %>
		<input type='radio' value='0' name='optxFuncao' <%=AUX%>><font face='<% =session("Fonte") %>' size='2'>Não
		</tr>
		</table>
	</td>
	</tr>
	<tr>	
	<td width='100%'>
		<table border='0' width='100%' cellspacing='0' style='margin-top: 2px'>
		<tr>
		<td width="18%"><font face="<% =session("Fonte") %>" size="2">13o. Salário:</font></td>
		<td width="82%"><font face="<% =session("Fonte") %>" size="2"></font>
		<% if (x_Funcao and 64) = 64 then AUX = "checked" else AUX = "" %>
		<input type='radio' value='64' name='optxFuncao13' <%=AUX%> ><font face='<% =session("Fonte") %>' size='2'>Sim
		<% if (x_Funcao and 64) = 64 then AUX = "" else AUX = "checked" %>
		<input type='radio' value='0' name='optxFuncao13' <%=AUX%>><font face='<% =session("Fonte") %>' size='2'>Não
		</tr>
		</table>		
	</td>
	</tr>	
	<tr>	
	<td width='100%'>
		<table border='0' width='100%' cellspacing='0' style='margin-top: 2px'>
		<tr>
		<td width="18%"><font face="<% =session("Fonte") %>" size="2">IN 1.343:</font></td>
		<td width="82%"><font face="<% =session("Fonte") %>" size="2"></font>
		<% if (x_Funcao and 128) = 128 then AUX = "checked" else AUX = "" 
		%>
		<input type='radio' value='128' name='optxFuncaoIN' <%=AUX%> ><font face='<% =session("Fonte") %>' size='2'>Sim
		<% if (x_Funcao and 128) = 128 then AUX = "" else AUX = "checked" %>
		<input type='radio' value='0' name='optxFuncaoIN' <%=AUX%>><font face='<% =session("Fonte") %>' size='2'>Não
		</tr>
		</table>		
	</td>
	</tr>	
	<tr>	
	<td width='100%'>
		<table border='0' width='100%' cellspacing='0' style='margin-top: 2px'>
		<tr>
		<td width="18%"><font face="<% =session("Fonte") %>" size="2">Pecúlio por Invalidez:</font></td>
		<td width="82%"><font face="<% =session("Fonte") %>" size="2"></font>
		<% if (x_Funcao and 256) = 256 then AUX = "checked" else AUX = "" 		
		%>
		<input type='radio' value='256' name='optxFuncaoInv' <%=AUX%> ><font face='<% =session("Fonte") %>' size='2'>Sim
		<% if (x_Funcao and 256) = 256 then AUX = "" else AUX = "checked" %>
		<input type='radio' value='0' name='optxFuncaoInv' <%=AUX%>><font face='<% =session("Fonte") %>' size='2'>Não
		</tr>
		</table>		
	</td>
	</tr>	

	<tr>	
	<td width='100%'>
		<table border='0' width='100%' cellspacing='0' style='margin-top: 2px'>
		<tr>
		<td width="18%"><font face="<% =session("Fonte") %>" size="2">Pecúlio por Morte:</font></td>
		<td width="82%"><font face="<% =session("Fonte") %>" size="2"></font>
		<% if (x_Funcao and 512) = 512 then AUX = "checked" else AUX = "" 
		%>
		<input type='radio' value='512' name='optxFuncaoMorte' <%=AUX%> ><font face='<% =session("Fonte") %>' size='2'>Sim
		<% if (x_Funcao and 512) = 512 then AUX = "" else AUX = "checked" %>
		<input type='radio' value='0' name='optxFuncaoMorte' <%=AUX%>><font face='<% =session("Fonte") %>' size='2'>Não
		</tr>
		</table>		
	</td>
	</tr>	
</table>
	


<table border='0' width='100%' style='border: 1px solid rgb(0,0,0)' cellspacing='0' style='margin-top:10px'>
<tr>
<td width='100%'>
	<table border='0' width='100%' cellspacing=2 cellpadding=0>
	<tr>
	<% if b_Serv = "1" then Aux = "checked" else Aux = "" %>	
	<td width='100%'><font face='<% =session("Fonte") %>' size='2'><input type='checkbox' name='chkbServ' value='1' <%=Aux%>>Evento de Serviço</font></td>
	</tr>
	</table>
</td>
</tr>	

<%
Aux = "display:none"
if t_Evento = "1" then Aux = ""
%>
<tr>
<td width='100%'>	
	<table border='0' width='100%' cellspacing=2 cellpadding=0 id="vis_DP" style="<%=Aux%>">	
	<tr>
	<td width='100%'><font face='<% =session("Fonte") %>' size='2'><input type='checkbox' name='chkxFuncao' value='1' <%=FN(1)%>>Data da provisão igual a data de emissão do documento para interface</font></td>
	</tr>
	</table>
</td>
</tr>	
</table>


<table border='0' width='100%' style='border: 1px solid rgb(0,0,0)' cellspacing='0' style='margin-top:5px'>
<tr>
<td width="100%">
	<table border="0" width="100%" cellspacing="0">        
	<tr>
	<td width="18%"><font face="<% =session("Fonte") %>" size="2">Tipo Doc: </font></td>
	<td width="82%">
    <select name="LstTDocFiscal" size="1" style="font-family: <% =session("Fonte") %>; font-size: 8pt; width: 250">
    <option value="" selected></option>
<%	
        
    SQL = "exec spr_lst_t_doc_fiscal "
    SQL = SQL & "0"          'n_Nat_Serv
    SQL = SQL & ",''"        'Par
    SQL = SQL & ",1"       'nAcao
    
                    
    Set oRsTDF = DbObj.Execute(SQL)

    if not oRsTDF.eof then
       
		While not oRsTDF.eof 
					  
			if Trim(oRsTDF("t_Doc_Fiscal")) = t_Doc_Fiscal then
			%>
				<option value="<% =oRsTDF("t_Doc_Fiscal") %>" selected><% =oRsTDF("Dscr") %>
			<%
			else
			%>
			  <option value="<% =oRsTDF("t_Doc_Fiscal") %>"><% =oRsTDF("Dscr") %>
			<%  
			end if
		
			oRsTDF.movenext
		Wend

    end if

%>
	</select>
	</td>
	</tr>			
	<tr>
	<td width="18%"><font face="<% =session("Fonte") %>" size="2">Serviço: </font></td>
	<td width="82%">
    <select name="LstNatServ" size="1" style="font-family: <% =session("Fonte") %>; font-size: 8pt; width: 370">
    <option value="" selected></option>
<%	
        
    SQL = "exec spr_lst_nat_serv "
    SQL = SQL & "0"          'n_Nat_Serv
    SQL = SQL & ",''"        'Par
    SQL = SQL & ",100"       'nAcao
    SQL = SQL & "," & n_Emp
    SQL = SQL & ",-1"        'n_Filial
    SQL = SQL & ",0"         'i_Cad
    SQL = SQL & ",0"         't_Documento
    SQL = SQL & ",0"         'd_Venc
    SQL = SQL & ",0"         'i_Evento
                        
    Set oRsNS = DbObj.Execute(SQL)

    if not oRsNS.eof then
       
		While not oRsNS.eof 
					  
			if Trim(oRsNS("n_Nat_Serv")) = n_Nat_Serv then
			%>
				<option value="<% =oRsNS("n_Nat_Serv") %>" selected><% =oRsNS("Dscr1") %>
			<%
			else
			%>
			  <option value="<% =oRsNS("n_Nat_Serv") %>"><% =oRsNS("Dscr1") %>
			<%  
			end if
		
			oRsNS.movenext
		Wend

    end if

%>
	</select>
	<img src="../../Figuras/Ajuda.gif" ONCLICK="NSHelp" alt="Ajuda" WIDTH="21" HEIGHT="23">	
	</td>
	</tr>		
	</table>
	
</td>
</tr>		
</table>
	

		
	<script language="vbscript">
	
Sub NSHelp

	 Dim Abre, ende

	 if ManEve.LstNatServ.value <> "" then

		if isnumeric(ManEve.LstNatServ.value) then
		

			Help ManEve.LstNatServ.value ,"2"

		 end if

	 end if

End Sub
	
	
	<% if b_Exib_Trib_Evento = 0 then %>
				ManEve.optbCalcIRRF.item(0).checked = true
				ManEve.optbCalcISS.item(0).checked = true
				ManEve.optbCalcINSS.item(0).checked = true

				<% if b_Calc_IRRF = 0 then %>
					ManEve.optbCalcIRRF.item(1).checked = true 
				<% end if %>
				
				<% if b_Calc_IRRF = 5 then %>
					ManEve.optbCalcIRRF.item(2).checked = true 
				<% end if %>

				<% if b_Calc_IRRF = 6 then %>
					ManEve.optbCalcIRRF.item(3).checked = true 
				<% end if %>
					
				<% if b_Calc_ISS = 0 then %>
					ManEve.optbCalcISS.item(1).checked = true
				<% end if%>

				<% if b_Calc_ISS = 5 then %>
					ManEve.optbCalcISS.item(2).checked = true
				<% end if%>
				

				<% if b_Calc_INSS = 0 then %>
					ManEve.optbCalcINSS.item(1).checked = true
				<% end if%>
				
				<% if b_Calc_INSS = 5 then %>
					ManEve.optbCalcINSS.item(2).checked = true
				<% end if%>

			<% else %>
			
				ManEve.chkbCalcIRRFPF.checked = false
				ManEve.chkbCalcIRRFPJ.checked = false
				ManEve.chkbCalcISSPF.checked = false
				ManEve.chkbCalcISSPJ.checked = false
				ManEve.chkbCalcINSSPF.checked = false
				ManEve.chkbCalcINSSPJ.checked = false
			

				<% if b_Calc_IRRF = 2 then %>
					ManEve.chkbCalcIRRFPF.checked = true 
				<% end if %>
					
				<% if b_Calc_IRRF = 3 then %>
					ManEve.chkbCalcIRRFPJ.checked = true 
				<% end if %>
				
				<% if b_Calc_IRRF = 4 then %>
					ManEve.chkbCalcIRRFPF.checked = true
					ManEve.chkbCalcIRRFPJ.checked = true 					 
				<% end if %>

				<% if b_Calc_IRRF = 5 then %>
					ManEve.chkbCalcIRRFInter.checked = true
				<% end if%>	
				
				<% if b_Calc_IRRF = 6 then %>
					ManEve.chkbCalcIRRFRetInter.checked = true
				<% end if%>	
				
				
				<% if b_Calc_ISS = 2 then %>
					ManEve.chkbCalcISSPF.checked = true 
				<% end if %>
					
				<% if b_Calc_ISS = 3 then %>
					ManEve.chkbCalcISSPJ.checked = true 
				<% end if %>
				
				<% if b_Calc_ISS = 4 then %>
					ManEve.chkbCalcISSPF.checked = true
					ManEve.chkbCalcISSPJ.checked = true  
				<% end if %>
				
				<% if b_Calc_ISS = 5 then %>
					ManEve.chkbCalcISSInter.checked = true
				<% end if%>	

				<% if b_Calc_ISS = 6 then %>
					ManEve.chkbCalcISSRetInter.checked = true
				<% end if%>	

				<% if b_Calc_INSS = 2 then %>
					ManEve.chkbCalcINSSPF.checked = true 
				<% end if %>
					
				<% if b_Calc_INSS = 3 then %>
					ManEve.chkbCalcINSSPJ.checked = true 
				<% end if %>
				
				<% if b_Calc_INSS = 4 then %>
					ManEve.chkbCalcINSSPF.checked = true
					ManEve.chkbCalcINSSPJ.checked = true  
				<% end if %>

				<% if b_Calc_INSS = 5 then %>
					ManEve.chkbCalcINSSInter.checked = true
				<% end if%>	
			
				<% if b_Calc_INSS = 6 then %>
					ManEve.chkbCalcINSSRetInter.checked = true
				<% end if%>	

			<% end if %>
			
</script>


<script language="vbscript">

Sub cbonNatRendPF_onChange

	if ManEve.cbonNatRendPF.value = "561" then
		DivInfRendPF.style.display = "inline"
	else
		DivInfRendPF.style.display = "none"
	end if
		
End Sub
	Function ExisteCampo(Campo)
	
		dim i
				
		ExisteCampo = 0
		For i = 0 to ManEve.elements.length - 1
			
			if ManEve.elements(i).name = Campo then 
				ExisteCampo = 1
				exit for
			end if
			
		next
		
	End Function


	sub ExibNat()
		if ManEve.optbCalcIRRF(0).checked or ManEve.optbCalcIRRF(2).checked then
			DivNatRendPF.style.display = "inline"
			DivNatRendPJ.style.display = "inline"
			if ManEve.cbonNatRendPF.value = "561" then
				DivInfRendPF.style.display = "inline"
			else
				DivInfRendPF.style.display = "none"
			end if	
		else
			DivNatRendPF.style.display = "none"
			DivInfRendPF.style.display = "none"
			DivNatRendPJ.style.display = "none"
		end if


		if ManEve.optbCalcIRRF(0).checked or ManEve.optbCalcIRRF(2).checked then
			DivDoencaGrave.style.display = "inline"
		else
			DivDoencaGrave.style.display = "none"
		end if

	end sub
	Sub ATUTELA(v)
			
		if v="IRRF5" then 
			if ManEve.chkbCalcIRRFPF.checked = true then ManEve.chkbCalcIRRFPF.checked = false
			if ManEve.chkbCalcIRRFPJ.checked = true then ManEve.chkbCalcIRRFPJ.checked = false
		end if
	
		if v="IRRF6" then 
			if ManEve.chkbCalcIRRFPF.checked = true then ManEve.chkbCalcIRRFPF.checked = false
			if ManEve.chkbCalcIRRFPJ.checked = true then ManEve.chkbCalcIRRFPJ.checked = false
			if ManEve.chkbCalcIRRFInter.checked = true then ManEve.chkbCalcIRRFInter.checked = false
		end if
		
		if v="IRRF7" then
			if ManEve.chkbCalcIRRFPF.checked = true then ManEve.chkbCalcIRRFPF.checked = false
			if ManEve.chkbCalcIRRFPJ.checked = true then ManEve.chkbCalcIRRFPJ.checked = false
			if ManEve.chkbCalcIRRFRetInter.checked = true then ManEve.chkbCalcIRRFRetInter.checked = false
		end if

		if v="IRRF2" or v="IRRF3" then
			if ManEve.chkbCalcIRRFRetInter.checked = true then ManEve.chkbCalcIRRFRetInter.checked = false
			if ManEve.chkbCalcIRRFInter.checked = true then ManEve.chkbCalcIRRFInter.checked = false
		end if
		if left(v,4) = "IRRF" then
			if ManEve.chkbCalcIRRFPF.checked or ManEve.chkbCalcIRRFInter.checked  then
				DivNatRendPF.style.display = "inline"
				
				if ManEve.cbonNatRendPF.value = "561" then
					DivInfRendPF.style.display = "inline"
				else
					DivInfRendPF.style.display = "none"
				end if
			else
				DivNatRendPF.style.display = "none"
                DivInfRendPF.style.display = "none"
			end if
			if ManEve.chkbCalcIRRFPJ.checked or ManEve.chkbCalcIRRFInter.checked  then
				DivNatRendPJ.style.display = "inline"
			else
				DivNatRendPJ.style.display = "none"
			end if			
			
		end if
						
		if v="ISS2" or v="ISS3" then
			if ManEve.chkbCalcISSInter.checked = true then ManEve.chkbCalcISSInter.checked = false
			if ManEve.chkbServ.checked = false then ManEve.chkbServ.checked = true
		end if
		
		if v="ISS5" then 
			if ManEve.chkbCalcISSPF.checked = true then ManEve.chkbCalcISSPF.checked = false
			if ManEve.chkbCalcISSPJ.checked = true then ManEve.chkbCalcISSPJ.checked = false
		end if
		
		if v="ISS6" then 
			if ManEve.chkbCalcISSPF.checked = true then ManEve.chkbCalcISSPF.checked = false
			if ManEve.chkbCalcISSPJ.checked = true then ManEve.chkbCalcISSPJ.checked = false
			if ManEve.chkbCalcISSInter.checked = true then ManEve.chkbCalcISSInter.checked = false
		end if
		
	

		if v="INSS2" or v="INSS3" then
			if ManEve.chkbCalcINSSInter.checked = true then ManEve.chkbCalcINSSInter.checked = false
		end if
		
		if v="INSS5" then 
			if ManEve.chkbCalcINSSPF.checked = true then ManEve.chkbCalcINSSPF.checked = false
			if ManEve.chkbCalcINSSPJ.checked = true then ManEve.chkbCalcINSSPJ.checked = false
		end if


		if v="INSS6" then 
			if ManEve.chkbCalcINSSPF.checked = true then ManEve.chkbCalcINSSPF.checked = false
			if ManEve.chkbCalcINSSPJ.checked = true then ManEve.chkbCalcINSSPJ.checked = false
			if ManEve.chkbCalcINSSInter.checked = true then ManEve.chkbCalcINSSInter.checked = false
		end if

	End Sub

</script>


<table border="0" width="100%" style="margin-top: 15px" cellspacing="0">
  <tr>
    <td width="100%"><font face="<% =session("Fonte") %>" size="2">Descrição do Evento:</font></td>
  </tr>
  <tr>
    <td width="100%"><textarea rows="7" name="txtDscr" cols="62"></textarea></td>
  </tr>
</table>
<%
	oRs.close
	DbObj.close
	Set oRs = nothing
	Set DbObj = nothing
%>

<script language="vbscript">
	top.frames("rodape").location.href="RDPCTB.Asp?Tabela=81"
</script>
<script language="vbscript">


Sub optbCalcISS_onClick

  if ManEve.optbCalcISS.item(0).checked = true then 
     if ManEve.chkbServ.checked = false then ManEve.chkbServ.checked = true 
  end if
     
End Sub 


<% if b_Exib_Cod_Trib = "1" then %>

	<% if b_Trib_Evento = "0" then %>
		document.all("cboiTribIRPJ").disabled = "True"
		document.all("cboiTribIRPF").disabled = "True"
	<% else %>
		document.all("cboiTribIRPJ").disabled = "False"
		document.all("cboiTribIRPF").disabled = "False"
	<% end if %>
<%
 end if
%>	

Sub cbobTribEvento_onChange

	if ManEve.cbobTribEvento.value = "0" then 
		ManEve.cboiTribIRPJ.value = "0"
		ManEve.cboiTribIRPF.value = "0"
		document.all("cboiTribIRPJ").disabled = "True"
		document.all("cboiTribIRPF").disabled = "True"
	else
		document.all("cboiTribIRPJ").disabled = "False"
		document.all("cboiTribIRPF").disabled = "False"
	end if	
		
End Sub

	
		

	ManEve.txtDscr.value = "<% =Dscr %>"

	'ManEve.txtEvento.focus()
 
		
	Sub cbotEvento_onChange

		if ManEve.cbotEvento.value = "1" then vis_DP.style.display = "" else vis_DP.style.display = "none"
		
		ManEve.elements(1).focus()
		
	End Sub


Sub window_onLoad()
<%if b_Exib_Trib_Evento = "1" then%>
	ATUTELA("IRRF")	
<%else%>
	ExibNat()
<%end if%>
end Sub

</script>
</body>
</html>
