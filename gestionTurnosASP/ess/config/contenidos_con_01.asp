<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: berths_con_01.asp
'Descripción: ABM de Berths
'Autor : Raul Chinestra
'Fecha: 23/11/2007

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_buqnro

l_filtro = request("filtro")
l_orden  = request("orden")

l_buqnro  = request("buqnro")

if l_orden = "" then
  l_orden = " ORDER BY merdes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%> Contenidos - Buques</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" method="post">
<input type="hidden" name="buqnro" value="<%= l_buqnro %>">

<table>
    <tr>
        <th>Mercadería</th>
        <th>Exportadora</th>		
        <th>Sitio</th>				
	    <th>Toneladas</th>						
	    <th>Destino</th>		
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT connro, buq_mercaderia.mernro,merdes, expdes, sitdes, conton, desdes "
l_sql = l_sql & " FROM buq_contenido "
l_sql = l_sql & " INNER JOIN buq_mercaderia ON buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " INNER JOIN buq_exportadora ON buq_exportadora.expnro = buq_contenido.expnro "
l_sql = l_sql & " INNER JOIN buq_sitio ON buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " LEFT JOIN buq_destino ON buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " WHERE buq_contenido.buqnro = " & l_buqnro
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="5">No existen Contenidos cargados.</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('contenidos_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value + '&buqnro=<%= l_buqnro %>','',480,260)" onclick="Javascript:Seleccionar(this,<%= l_rs("connro")%>)">
	        <td width="20%" nowrap><%= l_rs("merdes")%></td>
	        <td width="20%" nowrap><%= l_rs("expdes")%></td>			
	        <td width="20%" nowrap><%= l_rs("sitdes")%></td>						
	        <td width="20%" align="right" nowrap><%= l_rs("conton")%></td>			
	        <td width="20%" nowrap><%= l_rs("desdes")%></td>						
	    </tr>
	<%
		l_rs.MoveNext
	loop
end if
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>

<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
