<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

Dim l_tipo
Dim l_rs
Dim l_sql
Dim l_id
Dim l_dni
Dim l_nrohistoriaclinica
Dim texto

texto = ""
l_tipo		    = request.Form("tipo")
l_id            = request.Form("id")
l_dni 	= request.Form("dni")
l_nrohistoriaclinica = request.Form("nrohistoriaclinica")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

texto = "OK"	
if l_dni <> "" then

	'Verifico que no este repetido el DNI
	l_sql = "SELECT * "
	l_sql = l_sql & " FROM clientespacientes "
	l_sql = l_sql & " WHERE dni=" & l_dni 
	if l_tipo = "M" then
		l_sql = l_sql & " AND id <> " & l_id
	end if
	l_sql  = l_sql  & " AND clientespacientes.empnro = " & Session("empnro")	
	
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
	    texto =  "Ya existe otro Paciente con el Nro de DNI ingresado."
	else
		texto = "OK"
	end if 		
	l_rs.close
end if	

if texto = "OK" then
	if l_nrohistoriaclinica <> "0" then	
		'Verifico que no este repetida el nro de historia clinica
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM clientespacientes "
		l_sql = l_sql & " WHERE nrohistoriaclinica='" & l_nrohistoriaclinica & "'" 
		if l_tipo = "M" then
			l_sql = l_sql & " AND id <> " & l_id
		end if
		l_sql = l_sql & " and clientespacientes.empnro = " & Session("empnro")  
		'Response.write	l_sql
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then
			texto =  "Ya existe otro Paciente con el Nro de Historia Clinica ingresado."
		else
			texto = "OK"
		end if 		
		l_rs.close
	end if
		
end if	
	
%>

<% Response.write texto %>

<%
Set l_rs = Nothing 
%>

