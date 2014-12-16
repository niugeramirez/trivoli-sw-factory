<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!-----------------------------------------------------------------------------------------------
Archivo		: consulta_busquedas_eyp_01.asp
Descripción	: Ifrm que muestra las Búsquedas que no estan asociadas a un postulante
Autor 		: Lisandro Moro
Fecha		: 27/05/2004
-------------------------------------------------------------------------------------------------
-->
<% 
 Dim l_rs
 Dim l_sql
 Dim l_orden
 
 Dim l_ternro
 
 l_ternro = request.QueryString("ternro")
 
 if l_orden = "" then
 	l_orden = " ORDER BY pos_busqueda.busnro "
 end if
 
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>B&uacute;squedas - Empleos y Postulantes - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,texto){
	if (jsSelRow != null)
		Deseleccionar(jsSelRow);

	document.datos.cabnro.value = cabnro;
	fila.className = "SelectedRow";
	jsSelRow		= fila;
	//Reemplazo los caracteres que agregue al crear la tabla
	var r;
	r = texto.replace(/~/gi,"'");
	parent.document.all.textsql.value = r ;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<table>
    <tr>
        <th>C&oacute;d.</th>
        <th>Descripción</th>
        <th>Requerimiento</th>
    </tr>
<%
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = " SELECT DISTINCT pos_busqueda.busnro, pos_busqueda.busdesabr, pos_busqueda.busnro, pos_busqueda.busfecha, pos_busqueda.busformal , pos_reqbus.reqbusnro, buscarac, reqpercarac, pos_reqpersonal.reqperdesabr "
l_sql = l_sql & " FROM pos_reqbus "
l_sql = l_sql & " INNER JOIN pos_busqueda ON pos_reqbus.busnro = pos_busqueda.busnro "
l_sql = l_sql & " LEFT JOIN pos_reqpersonal ON pos_reqpersonal.reqpernro = pos_reqbus.reqpernro "
l_sql = l_sql & " WHERE pos_reqbus.reqbusnro NOT IN( "
l_sql = l_sql & " 	Select pos_terreqbus.reqbusnro "
l_sql = l_sql & " 	FROM pos_terreqbus "
l_sql = l_sql & " 	where ternro = " & l_ternro & " ) "
 
 l_sql = l_sql & l_orden
 
 rsOpen l_rs, cn, l_sql, 0 
 if l_rs.eof then
 	%>
	<tr>
		<td colspan="4">No existen B&uacute;squedas</td>
	</tr>
	<%
 else
	do until l_rs.eof
		%>
	<% If l_rs("busformal") then %>
    	<tr onclick="Javascript:Seleccionar(this,<%= l_rs("reqbusnro")%>,'<%= replace(l_rs("reqpercarac"),"'","~")%>');">
	<% Else  %>
		<tr onclick="Javascript:Seleccionar(this,<%= l_rs("reqbusnro")%>,'<%= replace(l_rs("buscarac"),"'","~") %>');">
	<% End If %>
        	<td align="left"><%= l_rs("reqbusnro")%></td>
        	<td><%= l_rs("busdesabr")%></td>
			<% If l_rs("busformal") then %>
				<td align="left"><%= l_rs("reqperdesabr")%></td>
			<% Else  %>
				<td align="left">Informal</td>
			<% End If %>
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
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="">
</form>
</body>
</html>
