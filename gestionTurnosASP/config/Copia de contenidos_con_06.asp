<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: berths_con_06.asp
'Descripción: ABM de Berths
'Autor : Raul Chinestra
'Fecha: 23/11/2007

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_bernro
Dim l_berdes

Dim texto

texto = ""
l_tipo		= request.QueryString("tipo")
l_bernro    = request.QueryString("bernro")
l_berdes 	= request.QueryString("berdes")
'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT berdes"
l_sql = l_sql & " FROM for_berth "
l_sql = l_sql & " WHERE berdes='" & l_berdes & "'"
if l_tipo = "M" then
	l_sql = l_sql & " AND bernro <> " & l_bernro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otro Berth con esa Descripción."
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

