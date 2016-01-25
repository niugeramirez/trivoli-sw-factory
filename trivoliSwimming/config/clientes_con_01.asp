<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: clientes_con_01.asp
'Descripción: Grilla Administración de Clientes
'Autor : Trivoli
'Fecha: 31/05/2015

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
  l_orden = " ORDER BY clientes.nombre "
end if
%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<head>
<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Administracion de Clientes</title>
</head>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Nombre</th>
        <th>Tel</th>	
        <th>Celular</th>		
        <th>Mail</th>
		<th>Direccion</th>
		<th>Ciudad</th> 	
		<th>Acciones</th>		
    </tr>
    <%
    l_filtro = replace (l_filtro, "*", "%")

    Set l_rs = Server.CreateObject("ADODB.RecordSet")
    l_sql = "SELECT    clientes.* , ciudades.ciudad "
    l_sql = l_sql & " FROM clientes "
    l_sql = l_sql & " LEFT JOIN ciudades ON ciudades.id = clientes.idciudad "
	' Multiempresa
	if l_filtro <> "" then
	  l_sql = l_sql & " WHERE " & l_filtro & " "
	  l_sql = l_sql & " and clientes.empnro = " & Session("empnro")   
	else
		l_sql = l_sql & " where clientes.empnro = " & Session("empnro")   
	end if
	
	
	
    l_sql = l_sql & " " & l_orden

    rsOpen l_rs, cn, l_sql, 0 
    if l_rs.eof then
	    l_primero = 0
    %>
    
    <tr>
	    <td colspan="4" >No existen Clientes cargados para el filtro ingresado.</td>
    </tr>
    <%
    else
        l_primero = l_rs("id")
	    l_cant = 0
	    do until l_rs.eof
		    l_cant = l_cant + 1
	    %>
	    <tr ondblclick="Javascript:parent.abrirDialogo('dialog','clientes_con_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,650,350);" onclick="Javascript:parent.Seleccionar(this,<%= l_rs("id")%>,document.detalle_01.cabnro)">    
			<td width="10%" nowrap><%= l_rs("nombre")%></td>
			<td width="10%" align="center" nowrap><%= l_rs("telefono")%></td>
			<td width="10%" align="center" nowrap><%= l_rs("celular")%></td>
			<td width="10%" align="left" nowrap><%= l_rs("mail")%></td>			
	        <td width="10%" nowrap align="left"><%= l_rs("direccion")%></td>			
			<td width="10%" nowrap align="left"><%= l_rs("ciudad")%></td>				
	        <td align="center" width="10%" nowrap>                    
                <a href="Javascript:parent.abrirDialogo('dialog','clientes_con_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,650,250);"><img src="../shared/images/Modificar_16.png" border="0" title="Editar"></a>				                																												
				<a href="Javascript:parent.eliminarRegistroAJAX(document.detalle_01.cabnro,'dialogAlert','dialogConfirmDelete');"><img src="../shared/images/Eliminar_16.png" border="0" title="Baja"></a>
			</td>
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
<form name="detalle_01" id="detalle_01" method="post">
    <input type="hidden" id="cabnro" name="cabnro" value="0">
    <input type="hidden" name="orden" value="<%= l_orden %>">
    <input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
