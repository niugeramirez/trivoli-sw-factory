<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 



Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_id
Dim l_descripcion


Dim texto

texto = ""
l_tipo		    = request.Form("tipo")
l_id            = request.Form("id")
l_descripcion 	= request.Form("descripcion")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT descripcion"
l_sql = l_sql & " FROM recursosreservables "
l_sql = l_sql & " WHERE descripcion='" & l_descripcion & "'"
l_sql = l_sql & " and recursosreservables.empnro = " & Session("empnro")   
if l_tipo = "M" then
	l_sql = l_sql & " AND id <> " & l_id
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otro M&eacute;dico con ese apellido."
else
	texto = "OK"
end if 
l_rs.close
%>

<% Response.write texto %>

<%
Set l_rs = Nothing
%>

