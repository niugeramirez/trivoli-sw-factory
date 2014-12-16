<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: ag_solicitud_eventos_cap_06.asp
Descripción: Abm de Solicitud de Eventos. Verificaciones contra la BD
Autor : Raul Chinestra	
Fecha: 30/03/2004
-->
<% 

'on error goto 0

Dim l_tipo
Dim l_rs
Dim l_sql
Dim l_solnro
Dim l_soldesabr
Dim texto

texto = ""
l_tipo 		= request.QueryString("tipo")
l_solnro	= request.QueryString("solnro")
l_soldesabr = request.QueryString("soldesabr")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción
l_sql = "SELECT solnro "
l_sql = l_sql & " FROM cap_solicitud "
l_sql = l_sql & " WHERE soldesabr='" & l_soldesabr & "'"
if l_tipo = "M" then
	l_sql = l_sql & " AND solnro <> " & l_solnro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_rs.close
    texto =  "Existe otra Solicitud con esa Descripción."
    set l_rs = nothing
end if 

%>

<script>
<% 
 if texto <> "" then
%>
   parent.invalido('<% =texto %>')
<% else%>
   parent.valido();
<% end if%>
</script>

<%
Set l_rs = Nothing
%>

