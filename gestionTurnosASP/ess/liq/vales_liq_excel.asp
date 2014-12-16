<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=vales.xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->


<!--
-----------------------------------------------------------------------------
Archivo        : vales_liq_01.asp
Descripcion    : Modulo que se encarga de listar los vales.
Creador        : Scarpa D.
Fecha Creacion : 09/01/2004
Modificacion   :
  27/01/2003 - Scarpa D. - Mostrar los vales liquidados
  23/04/2004 - Hoffman J. - Se agrego el doble click
  08/06/2004 - Alvaro Bayon - Formateo el monto con dos decimales. Alineación
    Modificado  : 12/09/2006 Raul Chinestra - se agregó Vales en Autogestión   
	                28/09/2006 Maximiliano Breglia - se saco v_empleado  
-----------------------------------------------------------------------------
-->
<% 
'on error goto 0

Dim l_rs
Dim l_sql

Dim l_filtro
Dim l_orden

Dim l_pliqnro
Dim l_tvalenro
dim l_ternro

l_filtro = request("filtro")
l_orden  = request("orden")

l_pliqnro  = request("pliqnro")
l_tvalenro = request("tvalenro")

if l_orden = "" then
  l_orden = " ORDER BY valnro ASC"  'orden por número asc
end if

l_ternro = l_ess_ternro

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title>Configuraci&oacute;n Vales - RHPro &reg;</title>
</head>
<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro,ttabnro)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value    = cabnro;
 document.datos.pronro.value    = fila.pronro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>C&oacute;digo</th>
        <th>Tipo&nbsp;Vale</th>				
        <th>Descripci&oacute;n</th>		
        <th>Monto</th>				
        <th>Fec.Pedido</th>				
        <th>Revisado</th>						
        <th>Liq.</th>						
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql =         " SELECT vales.*,terape, terape2, ternom, ternom2, empleg, pronro, tipovale.* "
l_sql = l_sql & " FROM vales "
l_sql = l_sql & " INNER JOIN tipovale ON tipovale.tvalenro = vales.tvalenro "
l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = vales.empleado "
l_sql = l_sql & " WHERE vales.empleado =  " & l_ternro

if l_tvalenro <> "" then
  l_sql = l_sql & " AND vales.tvalenro=" & l_tvalenro & " "
end if

if l_pliqnro <> "" then
  l_sql = l_sql & " AND pliqdto=" & l_pliqnro & " "
end if

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if

l_sql = l_sql & " " & l_orden

rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="8">No hay datos</td>
</tr>
<%else
	do until l_rs.eof
	%>
	<tr pronro="<%= l_rs("pronro")%>" onclick="Javascript:Seleccionar(this,<%=l_rs("valnro")%>)"
	ondblclick="Javascript:parent.abrirVentana('vales_liq_02.asp?Tipo=M&valnro=<%= l_rs("valnro") %> &pliqnro=<%= l_pliqnro %>&tvalenro=<%= l_rs("valnro") %>','',550,300);">
		<td align="right" width="5%"><%=l_rs("valnro")%></td>
		<td width="15%"><%=l_rs("tvaledesabr")%> </td>
		<td width="10%"><%=l_rs("valdesc")%> </td>		
		<td align="right" width="10%"><%=formatnumber(l_rs("valmonto"),2)%> </td>		
		<td align="right" width="10%"><%=l_rs("valfecped")%> </td>		
		<td align="center" width="10%"><% if l_rs("valrevis") then response.write "SI" else response.write "NO" end if %> </td>				
		<%if isNUll(l_rs("pronro")) then %>
		<td align="center" width="5%">NO</td>				
		<%else%>
		<td align="center" width="5%">SI</td>						
		<%end if%>		
	</tr>
	<%l_rs.MoveNext
	loop
end if 
l_rs.Close
cn.Close	
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="pronro" value="">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
