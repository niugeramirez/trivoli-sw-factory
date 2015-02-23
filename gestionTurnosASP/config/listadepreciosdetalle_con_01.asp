<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_01.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden


Dim l_idcab

l_filtro = request("filtro")
l_orden  = request("orden")

l_idcab = request("idcab")

'response.write l_idtemplatereserva

if l_orden = "" then
  l_orden = " ORDER BY descripcion"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Detalle Lista de Precios</title>
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

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" >
<table>
    <tr>
        <th>Practica</th>
		<th>Precio</th>
       	
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT listapreciosdetalle.id , descripcion, precio "
l_sql = l_sql & " FROM listapreciosdetalle "
l_sql = l_sql & " INNER JOIN practicas ON listapreciosdetalle.idpractica = practicas.id "
l_sql = l_sql & " WHERE idlistaprecioscabecera = " & l_idcab
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="2" >No existen Detalles de Lista de Precios cargados.</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('listadepreciosdetalle_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',520,200)" onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">
	        <td width="15%" align="left" nowrap><%= l_rs("descripcion")%></td>
	        <td width="15%" align="right" nowrap><%= formatnumber(l_rs("precio"),2)%></td>
	       <!--  <td width="85%" nowrap><%'= l_rs("descripcion")%></td>		 -->	
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
