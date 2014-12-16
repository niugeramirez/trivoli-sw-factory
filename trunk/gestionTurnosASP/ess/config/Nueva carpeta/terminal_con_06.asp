<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 

'Archivo: terminal_con_06.asp
'Descripción: ABM de Terminales
'Autor : Gustavo Manfrin
'Fecha: 19/04/2005

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_ternro
Dim l_tercod
Dim l_terdes
Dim l_terizq

Dim texto

texto = ""
l_tipo		= request.QueryString("tipo")
l_ternro    = request.QueryString("ternro")
l_tercod 	= request.QueryString("tercod")
l_terdes 	= request.QueryString("terdes")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT tercod"
l_sql = l_sql & " FROM tkt_terminal "
l_sql = l_sql & " WHERE tercod='" & l_tercod & "'"
if l_tipo = "M" then
	l_sql = l_sql & " AND ternro <> " & l_ternro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otra terminal con ese Código."
else
	l_rs.close
	l_sql = "SELECT terdesc"
	l_sql = l_sql & " FROM tkt_terminal "
	l_sql = l_sql & " WHERE terdesc='" & l_terdes & "'"
	if l_tipo = "M" then
		l_sql = l_sql & " AND ternro <> " & l_ternro
	end if
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
	    texto =  "Ya existe otra terminal con esa Descripción."
	end if
end if 
l_rs.close
%>

<script>
<% 
 if texto <> "" then
%>
   parent.invalido('<%= texto %>')
<% else%>
   parent.valido();
<% end if%>
</script>

<%
Set l_rs = Nothing
%>

