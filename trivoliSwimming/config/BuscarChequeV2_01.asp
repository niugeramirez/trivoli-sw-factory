<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: contracts_con_01.asp
'Descripción: ABM de Contracts
'Autor : Raul Chinestra
'Fecha: 28/11/2007

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_totvol
Dim l_saldo
Dim l_cobrado
Dim l_monto_venta
Dim l_cant

Dim l_primero

l_filtro = request("filtro")
l_orden  = request("orden")


if l_orden = "" then
  l_orden = " ORDER BY  fecha_emision,  numero "
end if


'l_ternro  = request("ternro")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Buscar Cheque</title>
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

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Numero</th>
        <th>Fecha Emision</th>		
		<th>Banco</th>
		<th>Importe</th>
					
    </tr>
<%
l_filtro = replace (l_filtro, "**", "%")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT    cheques.id, cheques.fecha_emision , cheques.numero , bancos.nombre_banco , cheques.importe  "
l_sql = l_sql & " FROM cheques "
l_sql = l_sql & " LEFT JOIN bancos ON bancos.id = cheques.id_banco "


if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
  l_sql  = l_sql  & " AND cheques.empnro = " & Session("empnro")
else  
  l_sql = l_sql & " where cheques.empnro = " & Session("empnro")   
end if

l_sql = l_sql & " " & l_orden

'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="7" >No existen Cheques cargados para el filtro ingresado.</td>
</tr>
<%else
    l_primero = l_rs("id")
	l_cant = 0
	do until l_rs.eof
		l_cant = l_cant + 1
	%>
	    <tr ondblclick="Javascript:parent.AsignarCheque('<%= l_rs("id")%>','<%= l_rs("fecha_emision")%>','<%= l_rs("numero")%>','<%= l_rs("nombre_banco")%>','<%= l_rs("importe")%>')" onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">

	        <td align="center" width="10%" nowrap><%= l_rs("numero")%></td>
	        <td align="center" width="10%" nowrap><%= l_rs("fecha_emision")%></td>		
			<td width="10%" align="center" nowrap><%= l_rs("nombre_banco") %></td>	
			<td width="10%" align="center" nowrap><%= l_rs("importe") %></td>

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
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
