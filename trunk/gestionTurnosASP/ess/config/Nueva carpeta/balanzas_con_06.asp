<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: balazas_con_06.asp
'Descripción: ABM de Balanzas
'Autor : Gustavo Manfrin
'Fecha: 27/04/2005

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_balnro
Dim l_balcod
Dim l_baldes

Dim texto

texto = ""
l_tipo		= request.QueryString("tipo")
l_balnro    = request.QueryString("balnro")
l_balcod 	= request.QueryString("balcod")
l_baldes 	= request.QueryString("baldes")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT balcod"
l_sql = l_sql & " FROM tkt_balanza "
l_sql = l_sql & " WHERE balcod='" & l_balcod & "'"
if l_tipo = "M" then
	l_sql = l_sql & " AND balnro <> " & l_balnro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otra balanza con ese Código."
else
	l_rs.close
	l_sql = "SELECT baldes"
	l_sql = l_sql & " FROM tkt_balanza "
	l_sql = l_sql & " WHERE baldes='" & l_baldes & "'"
	if l_tipo = "M" then
		l_sql = l_sql & " AND balnro <> " & l_balnro
	end if
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
	    texto =  "Ya existe otra balanza con esa Descripción."
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

