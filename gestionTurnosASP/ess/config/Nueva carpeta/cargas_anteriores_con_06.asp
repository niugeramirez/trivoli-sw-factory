<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: cargas_anteriores_con_06.asp
'Descripción: ABM de Cargas Anteriores
'Autor : Gustavo manfrin
'Fecha: 07/08/2006

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_lugnro
Dim l_pronro
Dim l_carconnro

Dim texto

texto = ""
l_tipo		= request.QueryString("tipo")
l_pronro    = request.QueryString("pronro")
l_carconnro	= request.QueryString("carconnro")
l_lugnro    = request.QueryString("lugnro")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT carconnro"
l_sql = l_sql & " FROM tkt_cargasconf "
l_sql = l_sql & " WHERE pronro=" & l_pronro 
l_sql = l_sql & " AND lugdesnro=" & l_lugnro 
if l_tipo = "M" then
	l_sql = l_sql & " AND carconnro <> " & l_carconnro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
   texto =  "Ya existen estos datos."
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

