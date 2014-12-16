<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'================================================================================
'Archivo		: cliente_eva_06.asp
'Descripción	: Validacion de descripcion unica
'Autor			: CCRossi
'Fecha			: 13-12-2004
'Modificado		:
'================================================================================

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_evaclinro
Dim l_evaclinom
Dim texto

texto = ""
l_tipo 		= request.QueryString("tipo")
l_evaclinro	= request.QueryString("evaclinro")
l_evaclinom	= request.QueryString("evaclinom")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción
l_sql = "SELECT evaclinro "
l_sql = l_sql & " FROM evacliente "
l_sql = l_sql & " WHERE evaclinom='" & l_evaclinom & "'"
if l_tipo = "M" then
	l_sql = l_sql & " AND evaclinro <> " & l_evaclinro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_rs.close
    texto =  "Existe otro Cliente con este Nombre."
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

