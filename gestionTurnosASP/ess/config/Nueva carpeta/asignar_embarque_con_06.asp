<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: asignar_embarque_con_06.asp
'Descripción: ABM de Asignación de embarque
'Autor : Gustavo Manfrin
'Fecha: 20/09/2006

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_asiembnro
Dim l_embnro
Dim l_tarcod
Dim l_camnro
Dim l_tranro

Dim texto

texto = ""
l_tipo		= request.QueryString("tipo")
l_asiembnro = request.QueryString("asiembnro")
l_embnro    = request.QueryString("embnro")
l_tarcod 	= request.QueryString("tarcod")
l_tranro 	= request.QueryString("tranro")
l_camnro 	= request.QueryString("camnro")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT *"
l_sql = l_sql & " FROM tkt_asiemb "
l_sql = l_sql & " WHERE tarcod='" & l_tarcod & "'"
'l_sql = l_sql & " AND ordnro=" & l_ordnro
if l_tipo = "M" then
	l_sql = l_sql & " AND asiembnro <> " & l_asiembnro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otro Camionero con ese Nro. de Tarjeta."
else
	l_rs.close
	l_sql = "SELECT tkt_asiemb.camnro "
	l_sql = l_sql & " FROM tkt_asiemb "
	l_sql = l_sql & " WHERE tkt_asiemb.camnro=" & l_camnro
	l_sql = l_sql & " AND tkt_asiemb.embnro=" & l_embnro
	if l_tipo = "M" then
		l_sql = l_sql & " AND asiembnro <> " & l_asiembnro
	end if
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
	    texto =  "El Camionero ya tiene un Nro. de Tarjeta para el Embarque seleccionado."
	else
		l_rs.close
		l_sql = "SELECT tkt_ord_tra.tranro "
		l_sql = l_sql & " FROM tkt_ord_tra "
		l_sql = l_sql & " INNER JOIN tkt_embarque ON tkt_embarque.ordnro = tkt_ord_tra.ordnro "
		l_sql = l_sql & " WHERE tkt_embarque.embnro = " & l_embnro
		l_sql = l_sql & " AND tkt_ord_tra.tranro=" & l_tranro
		rsOpen l_rs, cn, l_sql, 0
		if l_rs.eof then
		    texto =  "El Transportista no pertenece a la Orden de Trabajo del Embarque seleccionado."
		else
			l_rs.close
			l_sql = "SELECT tkt_cam_tra.camnro "
			l_sql = l_sql & " FROM tkt_cam_tra "
			l_sql = l_sql & " WHERE tkt_cam_tra.camnro = " & l_camnro
			l_sql = l_sql & " AND tkt_cam_tra.tranro=" & l_tranro
			rsOpen l_rs, cn, l_sql, 0
			if l_rs.eof then
			    texto =  "El Camionero no esta asociado al Transportista."
			end if
		end if
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

