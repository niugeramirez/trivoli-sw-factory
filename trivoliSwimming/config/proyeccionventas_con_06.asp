<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<% 



Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_id

dim l_fecha_desde 
dim l_fecha_hasta 
dim l_idconceptoCompraVenta 

Dim texto

texto = ""
l_tipo		    = request.Form("tipo")
l_id            = request.Form("id")

l_fecha_desde 			= request.Form("fecha_desde")
l_fecha_hasta 			= request.Form("fecha_hasta")
l_idconceptoCompraVenta = request.Form("idconceptoCompraVenta")


if len(l_fecha_desde) = 0 then
	l_fecha_desde = "null"
else 
	l_fecha_desde = cambiafecha(l_fecha_desde,"YMD",true)	
end if 

if len(l_fecha_hasta) = 0 then
	l_fecha_hasta = "null"
else 
	l_fecha_hasta = cambiafecha(l_fecha_hasta,"YMD",true)	
end if 

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

texto = "OK"
'Verifico que no este repetido o incluido en otor intervalo, o se cruce con otro intervalo.
'Por el momento lo comento dada la complejidad de la logica de controlas los intervalos.
' l_sql = "SELECT * "
' l_sql = l_sql & " FROM proyeccionVentas "
' l_sql = l_sql & " WHERE CONVERT(VARCHAR(10), proyeccionventas.fecha_desde, 111) >=" & l_fecha_desde & ""
' l_sql = l_sql & " AND CONVERT(VARCHAR(10), proyeccionventas.fecha_hasta, 111) <=" & l_fecha_hasta & ""
' l_sql = l_sql & " AND proyeccionVentas.idconceptoCompraVenta = " & l_idconceptoCompraVenta & ""
' l_sql = l_sql & " and proyeccionVentas.empnro = " & Session("empnro")   
' if l_tipo = "M" then
	' l_sql = l_sql & " AND id <> " & l_id
' end if
' response.write l_sql & "<br>"

' rsOpen l_rs, cn, l_sql, 0
' if not l_rs.eof then
    ' texto =  "Ya existe otro Banco con ese Nombre."
' else
	' texto = "OK"
' end if 
' l_rs.close
%>

<% Response.write texto %>

<%
Set l_rs = Nothing
%>

