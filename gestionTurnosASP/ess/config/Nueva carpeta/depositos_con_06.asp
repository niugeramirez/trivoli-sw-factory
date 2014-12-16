<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: depositos_con_06.asp
'Descripción: ABM de Depósitos
'Autor : Alvaro Bayon
'Fecha: 11/02/2005

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_depnro
Dim l_depcod
Dim l_depdes

Dim texto

texto = ""
l_tipo		= request.QueryString("tipo")
l_depnro    = request.QueryString("depnro")
l_depcod 	= request.QueryString("depcod")
l_depdes 	= request.QueryString("depdes")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT depdes"
l_sql = l_sql & " FROM tkt_deposito "
l_sql = l_sql & " WHERE depdes='" & l_depdes & "'"
if l_tipo = "M" then
	l_sql = l_sql & " AND depnro <> " & l_depnro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otro Depósito con esa Descripción."
else
	l_rs.close
	l_sql = "SELECT depcod"
	l_sql = l_sql & " FROM tkt_deposito "
	l_sql = l_sql & " WHERE depcod='" & Trim(l_depcod) & "'"
	if l_tipo = "M" then
		l_sql = l_sql & " AND depnro <> " & l_depnro
	end if
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
	    texto =  "Ya existe otro Depósito con ese Código."
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

