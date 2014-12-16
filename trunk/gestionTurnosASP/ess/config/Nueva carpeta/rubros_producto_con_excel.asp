<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Rubros por productos.xls" %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->

<% 
'Archivo: rubros_producto_con_excel.asp
'Descripción: Abm de rubros por productos
'Autor : Gustavo Manfrin
'Fecha: 17/03/2005
'Modificado: 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden


l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY lugdes,prodes,rubdes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<head>
<!--<link href="/ticket/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Rubros por producto - Ticket</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,pronro,rubnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	document.datos.pronro.value = pronro;
	document.datos.rubnro.value = rubnro;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
		<th colspan="17">Rubros por Productos</th>
	</tr>
    <tr>
        <th align="center">Lugar</th>
        <th>Producto</th>		
        <th>Rubro</th>		
        <th>Base</th>		
        <th>Max. Sin Merma</th>		
        <th>Rango Desde</th>		
        <th>Rango Hasta</th>		
        <th>Dcto. Fraccion</th>				
        <th>Considera en Carga</th>				
		<th>Obligatorio en Carga</th>				
		<th>Muestra Camara en Carga</th>				
        <th>Considera en Descarga</th>				
		<th>Obligatorio en Descarga</th>				
		<th>Muestra Camara en Descarga</th>				
        <th>Considera en Transito</th>				
		<th>Obligatorio en Transito</th>				
		<th>Muestra Camara en Transito</th>				
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT tkt_rub_pro.lugnro, tkt_rub_pro.pronro, tkt_rub_pro.rubnro, bascam, tolmax, valrefdes, valrefhas, desporfra, "
l_sql = l_sql & " carga, transito, descarga, valcar, valdes, valtra, concar, contra, condes, "
l_sql = l_sql & " tkt_lugar.lugcod, tkt_producto.prodes, tkt_rubro.rubdes "
l_sql = l_sql & " FROM tkt_rub_pro "
l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_rub_pro.lugnro= tkt_lugar.lugnro "
l_sql = l_sql & " INNER JOIN tkt_producto ON tkt_rub_pro.pronro= tkt_producto.pronro "
l_sql = l_sql & " INNER JOIN tkt_rubro ON tkt_rub_pro.rubnro= tkt_rubro.rubnro "

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 

if l_rs.eof then%>
<tr>
	 <td colspan="7">No existen Rubros por producto</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td  align="center"><%= l_rs("lugcod")%></td>
	        <td  nowrap><%= l_rs("prodes")%></td>
	        <td  nowrap><%= l_rs("rubdes")%></td>			
	        <td  align="right" nowrap><%= l_rs("bascam")%></td>						
            <td  align="right" nowrap><%= l_rs("tolmax")%></td>			
	        <td  align="right" nowrap><%= l_rs("valrefdes")%></td>						
	        <td  align="right" nowrap><%= l_rs("valrefhas")%></td>						
	        <td  align="right" nowrap><%= l_rs("desporfra")%></td>						
	        <td  align="center" nowrap><%if l_rs("carga") = -1 then %>SI<% else %>NO<% End If %></td>						
	        <td  align="center" nowrap><%if l_rs("valcar") = -1 then %>SI<% else %>NO<% End If %></td>									
	        <td  align="center" nowrap><%if l_rs("concar") = -1 then %>SI<% else %>NO<% End If %></td>									
	        <td  align="center" nowrap><%if l_rs("descarga") = -1 then %>SI<% else %>NO<% End If %></td>									
	        <td  align="center" nowrap><%if l_rs("valdes") = -1 then %>SI<% else %>NO<% End If %></td>						
	        <td  align="center" nowrap><%if l_rs("condes") = -1 then %>SI<% else %>NO<% End If %></td>						
	        <td  align="center" nowrap><%if l_rs("transito") = -1 then %>SI<% else %>NO<% End If %></td>						
	        <td  align="center" nowrap><%if l_rs("valtra") = -1 then %>SI<% else %>NO<% End If %></td>						
	        <td  align="center" nowrap><%if l_rs("contra") = -1 then %>SI<% else %>NO<% End If %></td>						
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
<input type="hidden" name="pronro" value="0">
<input type="hidden" name="rubnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
