<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'Archivo: contracts_con_03.asp
'Descripción: ABM de Contracts
'Autor : Raul Chinestra
'Fecha: 28/11/2007

on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_buqnro
Dim l_buqdes
Dim l_tipopenro
Dim l_tipbuqnro
Dim l_agenro
Dim l_buqfecdes
Dim l_buqfechas
Dim l_buqton

l_tipo 		  = request.querystring("tipo")
l_buqnro      = request.Form("buqnro")
l_buqdes      = request.Form("buqdes")
l_tipopenro   = request.Form("tipopenro")
l_tipbuqnro   = request.Form("tipbuqnro")
l_agenro 	  = request.Form("agenro")
l_buqfecdes	  = request.Form("buqfecdes")
l_buqfechas	  = request.Form("buqfechas")
l_buqton	  = request.Form("buqton")


if len(l_buqfecdes) = 0 then
	l_buqfecdes = "null"
else 
	l_buqfecdes = cambiafechahora(l_buqfecdes,"YMD",true)	& " " & "10:00:00" 
end if 
if len(l_buqfechas) = 0 then
	l_buqfechas = "null"
else 
	l_buqfechas = cambiafecha(l_buqfechas,"YMD",true)	
end if 

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO buq_buque "
	l_sql = l_sql & " (buqdes,tipopenro,tipbuqnro,agenro, buqfecdes, buqfechas, buqton)"
	l_sql = l_sql & " VALUES ('" & l_buqdes & "'," & l_tipopenro & "," & l_tipbuqnro & "," & l_agenro & "," & l_buqfecdes & "," & l_buqfechas & "," & l_buqton & ")"
else
	l_sql = "UPDATE buq_buque "
	l_sql = l_sql & " SET buqdes     = '" & l_buqdes & "'"	
	l_sql = l_sql & "    ,tipopenro  =  " & l_tipopenro
	l_sql = l_sql & "    ,tipbuqnro  =  " & l_tipbuqnro
	l_sql = l_sql & "    ,agenro      = " & l_agenro
	l_sql = l_sql & "    ,buqfecdes   = " & l_buqfecdes
	l_sql = l_sql & "    ,buqfechas   = " & l_buqfechas	
	l_sql = l_sql & "    ,buqton      = " & l_buqton
	l_sql = l_sql & " WHERE buqnro = " & l_buqnro
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

