<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 



Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_id
Dim l_numero


Dim texto

texto = ""
l_tipo		    = request.Form("tipo")
l_id            = request.Form("id")
l_numero	 	= request.Form("numero")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT numero "
l_sql = l_sql & " FROM cheques "
l_sql = l_sql & " WHERE numero='" & l_numero & "'"
l_sql = l_sql & " and cheques.empnro = " & Session("empnro")   
if l_tipo = "M" then
	l_sql = l_sql & " AND id <> " & l_id
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otro Cheque con ese Numero."
else
	texto = "OK"
end if 
l_rs.close
%>

<% Response.write texto %>

<%
Set l_rs = Nothing
%>

