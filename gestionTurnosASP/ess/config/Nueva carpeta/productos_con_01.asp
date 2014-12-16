<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 

'Archivo: productos_con_01.asp
'Descripción: Abm de  Productos
'Autor : Lisandro Moro
'Fecha: 09/02/2005

'Modificada por: Javier Posadas
'Fecha: 05/04/2005
'Descripción: Se agregó la posibilidad de habilitar/deshabilitar Productos

'Modificada por: Raul Chinestra
'Fecha: 24/02/2006
'Descripción: Se ordenaron los productos numericamente por codigo y no alfabeticamente. 


on error goto 0
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_todos


l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY cast(procod as int) "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_sel_multiple.js"></script>
<head>
<link href="/ticket/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Productos - Ticket</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

/*function Seleccionar(fila,cabnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}*/

function SeleccionarFila(fila,cabnro){
	document.datos.cabnro.value = cabnro;
	
	if (fila.className == "SelectedRow")
   		jsSelRow = null;
	else
   		jsSelRow = fila;
  	
	Seleccionar(fila,cabnro);
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th align="center">Código</th>
        <th>Descripci&oacute;n</th>
		<th align="center">Habilitado</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT pronro, procod, prodes, proest "
l_sql = l_sql & " FROM tkt_producto "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro 
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No existen Productos</td>
</tr>
<%else
	l_todos = CStr(trim(l_rs("pronro")))
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('productos_con_02.asp?cabnro=' + datos.cabnro.value,'',520,380)" onclick="Javascript:SeleccionarFila(this,<%= l_rs("pronro")%>)">
	        <td width="20%" align="center"><%= l_rs("procod")%></td>
	        <td width="80%" nowrap><%= l_rs("prodes")%></td>
			<td width="30%" nowrap align="center"><% if l_rs("proest") then %>Si<% Else %>No<% End If %></td>
	    </tr>
	<%
		l_rs.MoveNext
		if not l_rs.EoF then
			l_todos = l_todos & "," & Trim(l_rs("pronro"))
		end if
	loop
end if
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>

<form name="datos" method="post">
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
<input type="hidden" name="listanro" value="">
<input type="hidden" name="listatodos" value="<%= l_todos%>">
</form>
</table>

<script>
  setearObjDatos(document.datos.listanro, document.datos.listatodos);
</script>
</body>
</html>
