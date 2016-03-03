<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 



Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_id
Dim l_titulo


Dim texto

texto = ""
l_tipo		    = request.Form("tipo")
l_id            = request.Form("id")
l_titulo 	= request.Form("titulo")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT titulo"
l_sql = l_sql & " FROM listaprecioscabecera "
l_sql = l_sql & " WHERE titulo='" & l_titulo & "'"
l_sql = l_sql & " and listaprecioscabecera.empnro = " & Session("empnro")   
if l_tipo = "M" then
	l_sql = l_sql & " AND id <> " & l_id
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otra lista con ese T&iacute;tulo."
else
	texto = "OK"
end if 
l_rs.close
%>

<% Response.write texto %>

<%
Set l_rs = Nothing
%>

