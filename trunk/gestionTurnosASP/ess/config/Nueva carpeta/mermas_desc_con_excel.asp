<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Localidades.xls" %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->

<% 
'Archivo: mermas_desc_con_excel.asp
'Descripción: Abm de Tabla de descuentos
'Autor : Gustavo Manfrin
'Fecha: 29/04/2005
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
  l_orden = " ORDER BY prodes, rubdes, desvodes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Tabla de descuentos - Ticket</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
        <th align="center" colspan="7">Tabla de descuentos</th>
    </tr>
    <tr>
        <th>CodProd</th>
        <th>Producto</th>
        <th>CodRubro</th>
        <th>Rubro</th>		
        <th>Desde</th>		
        <th>Hasta</th>		
        <th>Descuento</th>		    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT tkt_descuento.pronro,tkt_descuento.rubnro,tkt_descuento.desvodes,tkt_descuento.desvohas,tkt_descuento.desval, "
l_sql = l_sql & "tkt_producto.prodes,tkt_producto.procod,tkt_rubro.rubcod, tkt_rubro.rubdes "
l_sql = l_sql & " FROM tkt_descuento  "
l_sql = l_sql & " INNER JOIN tkt_producto ON tkt_descuento.pronro = tkt_producto.pronro  "
l_sql = l_sql & " INNER JOIN tkt_rubro ON tkt_descuento.rubnro = tkt_rubro.rubnro  "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="7">No existen Descuentos</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td width="10%" nowrap><%= l_rs("procod")%></td>
            <td width="40%" nowrap><%= l_rs("prodes")%></td>
	        <td width="10%" nowrap><%= l_rs("rubcod")%></td>
            <td width="30%" nowrap><%= l_rs("rubdes")%></td>
	        <td width="10%"nowrap align=right><%= l_rs("desvodes")%></td>			
	        <td width="10%"nowrap align=right><%= l_rs("desvohas")%></td>			
	        <td width="10%"nowrap align=right><%= l_rs("desval")%></td>			
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
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
