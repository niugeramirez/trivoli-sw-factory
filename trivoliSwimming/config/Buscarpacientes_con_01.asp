<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
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
Dim l_cant

Dim l_primero

l_filtro = request("filtro")
l_orden  = request("orden")


if l_orden = "" then
  l_orden = " ORDER BY apellido, nombre "
end if


'l_ternro  = request("ternro")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Buscar Pacientes</title>
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


function AsignarPaciente(id, apellido, nombre, nrohistoriaclinica, dni, domicilio, telefono, osid,  os){
//alert(apellido);
 parent.opener.EncontrePaciente(id, apellido, nombre, nrohistoriaclinica, dni, domicilio, telefono,osid,  os);
	//opener.parent.document.datos.apellido = 'ra';
	window.parent.close();
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
         <!-- <th>&nbsp;</th>
        <th>Legajo</th>	
        <th>&nbsp;</th>  	-->
        <th>Apellido</th>
        <th>Nombre</th>		
        <th>Nro. Hist. Cl&iacute;nica</th>		
        <th>DNI</th>		
		<th align="left">Domicilio</th>	
		<th align="left">Tel&eacute;fono</th>	
			
    </tr>
<%
l_filtro = replace (l_filtro, "*", "%")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT    clientespacientes.id id, clientespacientes.* , obrassociales.id osid , obrassociales.descripcion "
l_sql = l_sql & " FROM clientespacientes "
l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "


if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
  l_sql  = l_sql  & " AND clientespacientes.empnro = " & Session("empnro")
end if

if l_filtro = "" then
  l_sql  = l_sql  & " WHERE clientespacientes.empnro = " & Session("empnro")
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
	    <tr ondblclick="Javascript:AsignarPaciente('<%= l_rs("id")%>','<%= l_rs("apellido")%>','<%= l_rs("nombre")%>', '<%= l_rs("nrohistoriaclinica")%>', '<%= l_rs("dni")%>', '<%= l_rs("domicilio")%>' , '<%= l_rs("telefono")%>' , '<%= l_rs("osid")%>' , '<%= l_rs("descripcion")%>' )" onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">
	        <!--<td width="10%" nowrap><%'= l_rs("buqnro")%></td>		-->
			<!-- <td width="2%" nowrap><%'= l_cant %></td>
	        <td width="10%" nowrap><%'= l_rs("legpar1")%>-<%'= l_rs("legpar2")%>/<%'= l_rs("legpar3")%></td>			-->
	        <!-- <td width="1%" nowrap><a href="Javascript:parent.abrirVentana('Editarpacientes_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value ,'',600,250);"><img src="/turnos/shared/images/cal.gif" border="0" alt="Editar Paciente"></a> </td>  -->
	        <td width="10%" nowrap><%= l_rs("apellido")%></td>
	        <td width="10%" nowrap><%= l_rs("nombre")%></td>		
	        <td width="10%" align="center"><%= l_rs("nrohistoriaclinica")%></td>								
	        <td width="10%" nowrap><%= l_rs("dni")%></td>			
	        <td width="10%" nowrap><%= l_rs("domicilio")%></td>		
			<td width="10%" nowrap><%= l_rs("telefono")%></td>				
	         <!--<td width="10%" nowrap><%'= l_rs("prodes")%></td>  -->			
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
<!--
<script>    
	parent.parent.ActPasos(<%'= l_primero %>,"","MENU");
    parent.parent.datos.pasonro.value = <%'= l_primero %>;
</script>
-->
</table>
<form name="datos" method="post">
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
