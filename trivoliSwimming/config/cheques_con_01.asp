<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: cheques_con_01.asp
'Descripción: Grilla Administración de Cheques
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
  l_orden = " ORDER BY cheques.fecha_vencimiento  desc ,cheques.fecha_emision desc "
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
<title>Administracion de Cheques</title>
</head>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Numero</th>		
        <th>Fecha Emision</th>			
        <th>Fecha vencimiento</th>			
        <th>Banco</th>
		
		<th>Importe</th>
		<th>Emitido por Cliente</th>
		<th>Emisor Tercero</th>	
		<th>Emitido por Franquicia</th> 	
		<th>Validacion BCRA</th> 			
		<th>Estado</th> 
		<th>Acciones</th>		
    </tr>
    <%
    l_filtro = replace (l_filtro, "*", "%")

    Set l_rs = Server.CreateObject("ADODB.RecordSet")
    l_sql = "SELECT    cheques.id ,cheques.numero ,cheques.fecha_emision ,cheques.fecha_vencimiento ,cheques.id_banco ,cheques.importe "
	l_sql = l_sql & " ,cheques.flag_emitidopor_cliente ,cheques.emisor ,cheques.created_by ,cheques.creation_date ,cheques.last_updated_by "
	l_sql = l_sql & " 	,cheques.last_update_date ,cheques.empnro ,ISNULL(cheques.flag_propio,0) as flag_propio , cheques.validacion_bcra, bancos.nombre_banco "
	l_sql = l_sql & "   , case   "
	l_sql = l_sql & " 		when cheques.flag_propio = -1 then  "
	l_sql = l_sql & " 			case "
	l_sql = l_sql & " 				/*Chequeo que este asociado a una compra*/ "
	l_sql = l_sql & " 				when (	select COUNT(*) "
	l_sql = l_sql & " 						from	cajaMovimientos "
	l_sql = l_sql & " 						inner join tiposMovimientoCaja on  tiposMovimientoCaja.id = cajaMovimientos.id "
	l_sql = l_sql & " 						where	tiposMovimientoCaja.flagCompra = -1 "
	l_sql = l_sql & " 						and		cajaMovimientos.idcheque = cheques .id		 "				
	l_sql = l_sql & " 						) > 0 then 'ENTREGADO' "
	l_sql = l_sql & " else 'PENDIENTE ENTREGAR'  "
	l_sql = l_sql & " end "
	l_sql = l_sql & " else  "
	l_sql = l_sql & " 			case  "
	l_sql = l_sql & " 				when (	/*chequeo si esta asociado a una venta*/ "
	l_sql = l_sql & " 						select COUNT(*) "
	l_sql = l_sql & " 						from	cajaMovimientos "
	l_sql = l_sql & " 						inner join tiposMovimientoCaja on  tiposMovimientoCaja.id = cajaMovimientos.id "
	l_sql = l_sql & " 						where	tiposMovimientoCaja.flagVenta = -1 "
	l_sql = l_sql & " 						and		cajaMovimientos.idcheque = cheques .id	 "
	l_sql = l_sql & " 						) > 0 then "
	l_sql = l_sql & " 								case "
	l_sql = l_sql & " 									/*Chequeo que este asociado a una compra*/ "
	l_sql = l_sql & " 									when (	select COUNT(*) "
	l_sql = l_sql & " 											from	cajaMovimientos "
	l_sql = l_sql & " 											inner join tiposMovimientoCaja on  tiposMovimientoCaja.id = cajaMovimientos.id "
	l_sql = l_sql & " 											where	tiposMovimientoCaja.flagCompra = -1 "
	l_sql = l_sql & " 											and		cajaMovimientos.idcheque = cheques .id						 "
	l_sql = l_sql & " 											) > 0 then 'ENTREGADO' "
	l_sql = l_sql & " 									else 'PENDIENTE ENTREGAR'  "
	l_sql = l_sql & " 								end		 "						
	l_sql = l_sql & " 				else 'PENDIENTE ASOCIAR VENTA' "
	l_sql = l_sql & " 			end "
	l_sql = l_sql & " 	end  AS estado_cheque     "
	l_sql = l_sql & " FROM cheques "
    l_sql = l_sql & " LEFT JOIN bancos ON cheques.id_banco = bancos.id "
	' Multiempresa
	if l_filtro <> "" then
	  l_sql = l_sql & " WHERE " & l_filtro & " "
	  l_sql = l_sql & " and cheques.empnro = " & Session("empnro")   
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
	    <td colspan="4" >No existen Cheques cargados para el filtro ingresado.</td>
    </tr>
    <%
    else
        l_primero = l_rs("id")
	    l_cant = 0
	    do until l_rs.eof
		    l_cant = l_cant + 1
	    %>
	    <tr ondblclick="Javascript:parent.abrirDialogo('dialog','cheques_con_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,650,350);" onclick="Javascript:parent.Seleccionar(this,<%= l_rs("id")%>,document.detalle_01.cabnro)">    
			<td width="10%" nowrap><%= l_rs("numero")%></td>
			
			<td width="10%" align="center" nowrap><%= l_rs("fecha_emision")%></td>			
			<td width="10%" align="center" nowrap><%= l_rs("fecha_vencimiento")%></td>
			
			<td width="10%" align="left" nowrap><%= l_rs("nombre_banco")%></td>	
				
	        <td width="10%" nowrap align="center"><%= l_rs("importe")%></td>			
			
			<td width="10%" nowrap align="center"><% if l_rs("flag_emitidopor_cliente") = 0 then response.write "NO" else response.write "SI" end if %></td>	
			<td width="10%" nowrap align="left"><%= l_rs("emisor")%></td>			
			<td width="10%" nowrap align="center"><% if l_rs("flag_propio") = 0 then response.write "NO" else response.write "SI" end if %></td>
			<td width="10%" nowrap align="left"><%= UCase(Left(l_rs("validacion_bcra"),1)) & LCase(Mid(l_rs("validacion_bcra"),2))%></td>	
			<td width="10%" nowrap align="left"><%= l_rs("estado_cheque")%></td>						
	        <td align="center" width="10%" nowrap>                    
                <a href="Javascript:parent.abrirDialogo('dialog','cheques_con_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,650,350);"><img src="../shared/images/Modificar_16.png" border="0" title="Editar"></a>				                																												
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
