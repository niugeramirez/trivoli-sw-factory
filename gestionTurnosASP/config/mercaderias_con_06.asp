<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: companies_con_06.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_mernro
Dim l_merdes

Dim texto

texto = ""
l_tipo		= request.QueryString("tipo")
l_mernro 	= request.QueryString("mernro")
l_merdes	= request.QueryString("merdes")
'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT merdes "
l_sql = l_sql & " FROM buq_mercaderia "
l_sql = l_sql & " WHERE merdes ='" & l_merdes & "'"
if l_tipo = "M" then
	l_sql = l_sql & " AND mernro <> " & l_mernro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otra Mercadería con esa Descripción."
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

