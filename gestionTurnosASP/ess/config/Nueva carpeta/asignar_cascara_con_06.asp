<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: asignar_cascara_con_06.asp
'Descripción: ABM de Asignación de Cáscara
'Autor : Raul Chinestra
'Fecha: 11/05/2005

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_asicasnro
Dim l_ordnro
Dim l_tarnro
Dim l_camnro

Dim texto

texto = ""
l_tipo		= request.QueryString("tipo")
l_asicasnro = request.QueryString("asicasnro")
l_ordnro    = request.QueryString("ordnro")
l_tarnro 	= request.QueryString("tarnro")
l_camnro 	= request.QueryString("camnro")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT *"
l_sql = l_sql & " FROM tkt_asicas "
l_sql = l_sql & " WHERE tarnro='" & l_tarnro & "'"
'l_sql = l_sql & " AND ordnro=" & l_ordnro
if l_tipo = "M" then
	l_sql = l_sql & " AND asicasnro <> " & l_asicasnro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otro Camionero con ese Nro. de Tarjeta."
else
	l_rs.close
	l_sql = "SELECT *"
	l_sql = l_sql & " FROM tkt_asicas "
	l_sql = l_sql & " WHERE camnro=" & l_camnro
	l_sql = l_sql & " AND ordnro=" & l_ordnro
	if l_tipo = "M" then
		l_sql = l_sql & " AND asicasnro <> " & l_asicasnro
	end if
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
	    texto =  "El Camionero ya tiene un Nro. de Tarjeta para la Orden de Trabajo seleccionada."
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

