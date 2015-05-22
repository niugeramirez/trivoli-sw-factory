<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: pacientes_con_01.asp
'Descripción: ABM de Pacientes
'Autor : Raul Chinestra
'Fecha: 09/10/2014

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_cant

Dim l_primero

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY apellido, nombre "
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
<title>Pacientes</title>
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
         <!-- <th>&nbsp;</th>
        <th>Legajo</th>	
        <th>Fec. Ingreso</th>  -->	
        <th>Apellido</th>
        <th>Nombre</th>		
		<th>Nro. Hist. Cl&iacute;nica</th>
		<th>Obra Social</th>
        <th>DNI</th>		
		<th align="left">Domicilio</th>	
		<th align="left">Tel&eacute;fono</th>	
			
    </tr>
<%
l_filtro = replace (l_filtro, "*", "%")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  clientespacientes.id id,   clientespacientes.* , obrassociales.descripcion"
l_sql = l_sql & " FROM clientespacientes "
l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden

'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="7" >No existen Pacientes cargados para el filtro ingresado.</td>
</tr>
<%else
    l_primero = l_rs("id")
	l_cant = 0
	do until l_rs.eof
		l_cant = l_cant + 1
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('pacientes_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',750,280)" onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">
	        <td width="10%" nowrap><%= l_rs("apellido")%></td>				
	        <td width="10%" nowrap><%= l_rs("nombre")%></td>		
	        <td width="10%" align="center"><%= l_rs("nrohistoriaclinica")%></td>		
			 <td width="10%" nowrap><%= l_rs("descripcion")%></td>					
	        <td width="10%" nowrap><%= l_rs("dni")%></td>			
	        <td width="10%" nowrap><%= l_rs("domicilio")%></td>		
			<td width="10%" nowrap><%= l_rs("telefono")%></td>				
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
