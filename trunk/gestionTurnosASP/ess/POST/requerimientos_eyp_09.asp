<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: requerimientos_eyp_09.asp
Descripción: Abm de requerimientos
Autor : Martin Ferraro
Fecha: 07/05/2004
Modificado  : 12/09/2006 Raul Chinestra - se agregó Requerimientos de Personal en Autogestión   
-->
<% 

Dim l_tipo
Dim l_rs
Dim l_sql
Dim l_reqpernro
Dim l_reqperdesabr
Dim l_reqperdesext
Dim texto

texto = ""
l_tipo 			= request.querystring("tipo")
l_reqpernro 	= request.QueryString("reqpernro")
l_reqperdesext 	= request.QueryString("reqperdesext")
l_reqperdesabr 	= request.QueryString("reqperdesabr")



'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")
response.write l_reqpernro
response.write l_reqperdesabr

'Verifico que no este repetida la descripción
l_sql = "SELECT reqpernro "
l_sql = l_sql & " FROM pos_reqpersonal "
l_sql = l_sql & " WHERE reqperdesabr ='" & l_reqperdesabr & "'"
if l_tipo = "M" then
	l_sql = l_sql & " AND reqpernro <> " & l_reqpernro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_rs.close
    texto =  "Existe otro Requerimiento con esa Descripción Abreviada."
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

