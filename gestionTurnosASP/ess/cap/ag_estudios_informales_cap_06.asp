<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: ag_estudios_informales_cap_06.asp
Descripción: Abm de Estudios Informales. Verificaciones contra la BD
Autor : Lisandro Moro
Fecha: 29/03/2004
-->
<% 

Dim l_tipo
Dim l_rs
Dim l_sql
Dim l_estinfnro
Dim l_estinfdesabr
Dim texto
Dim l_tipcurnro

texto = ""
l_tipo 			= request.QueryString("tipo")
l_estinfnro		= request.QueryString("estinfnro")
l_estinfdesabr  = request.QueryString("estinfdesabr")
l_tipcurnro		= request.QueryString("tipcurnro")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción
l_sql = "SELECT estinfnro "
l_sql = l_sql & " FROM cap_estinformal "
l_sql = l_sql & " WHERE estinfdesabr='" & l_estinfdesabr & "'"
l_sql = l_sql & " and tipcurnro=" & l_tipcurnro
if l_tipo = "M" then
	l_sql = l_sql & " AND estinfnro <> " & l_estinfnro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_rs.close
    texto =  "Existe otro Estudio Informal con esa Descripción Abreviada para el Tipo de Curso seleccionado."
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

