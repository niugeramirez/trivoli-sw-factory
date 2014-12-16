<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: impresoras_con_06.asp
'Descripción: ABM de Impresoras
'Autor : Lisandro Moro
'Fecha: 26/09/2005
'Modificado: 

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_impnro
Dim l_impnom
Dim l_impnomcom
Dim l_impmat

Dim texto

texto = ""
l_tipo		= request.QueryString("tipo")
l_impnro    = request.QueryString("impnro")
l_impnom 	= request.QueryString("impnom")
l_impnomcom	= request.QueryString("impnomcom")
l_impmat 	= request.QueryString("impmat")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT impnro "
l_sql = l_sql & " FROM tkt_impresora "
l_sql = l_sql & " WHERE impnom ='" & l_impnom & "'"
if l_tipo = "M" then
	l_sql = l_sql & " AND impnro <> " & l_impnro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otra impresora con ese Nombre."
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

