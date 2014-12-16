<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Contracts.xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 
'Archivo: countries_con_excel.asp
'Descripción: Consulta de Countries
'Autor : Raul Chinestra
'Fecha: 23/11/2007

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY coudes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%> Contracts - buques</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>P/S</th>	
        <th>Ctr Number</th>
        <th>Date</th>		
		<th>Client</th>	
		<th>Quality</th>			
        <th>Volumen (Kgrs)</th>		
		<th>Port</th>				
        <th>Term</th>
		<th>Company</th>
		<th>Product</th>		
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT    connro "
l_sql = l_sql & " ,conpursal "
l_sql = l_sql & " ,ctrnum "
l_sql = l_sql & " ,confec "
l_sql = l_sql & " ,for_client.clinro, clidesabr "
l_sql = l_sql & " ,for_quality.quadesabr "
l_sql = l_sql & " ,conquantity "
l_sql = l_sql & " ,for_port.pornro, pordes "
l_sql = l_sql & " ,for_term.ternro, terdes "
l_sql = l_sql & " ,for_company.comnro, comdesabr "
l_sql = l_sql & " ,for_product.pronro, prodesabr "
l_sql = l_sql & " FROM for_contract "
l_sql = l_sql & " INNER JOIN for_term    ON for_contract.ternro = for_term.ternro "
l_sql = l_sql & " INNER JOIN for_company ON for_contract.comnro = for_company.comnro "
l_sql = l_sql & " INNER JOIN for_client  ON for_contract.clinro = for_client.clinro "
l_sql = l_sql & " INNER JOIN for_port    ON for_contract.pornro = for_port.pornro "
l_sql = l_sql & " INNER JOIN for_product ON for_contract.pronro = for_product.pronro "
l_sql = l_sql & " INNER JOIN for_quality ON for_contract.quanro = for_quality.quanro "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="10" >No existen Contracts cargados para el filtro ingresado.</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('contracts_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',780,580)" onclick="Javascript:Seleccionar(this,<%= l_rs("connro")%>)">
	        <td width="5%" align="center" nowrap><%= l_rs("conpursal")%></td>		
	        <td width="10%" nowrap><%= l_rs("ctrnum")%></td>
	        <td width="10%" nowrap><%= l_rs("confec")%></td>			
	        <td width="10%" nowrap><%= l_rs("clidesabr")%></td>
	        <td width="10%" nowrap><%= l_rs("quadesabr")%></td>										
	        <td width="10%" nowrap><%= l_rs("conquantity")%></td>			
	        <td width="10%" nowrap><%= l_rs("pordes")%></td>			
	        <td width="10%" nowrap><%= l_rs("terdes")%></td>			
	        <td width="10%" nowrap><%= l_rs("comdesabr")%></td>
	        <td width="10%" nowrap><%= l_rs("prodesabr")%></td>			
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
