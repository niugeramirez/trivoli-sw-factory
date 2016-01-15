<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: columnas_reporte_01.asp
Descripcion: Modulo que lista las columnas de un ConfRep.
Modificacion:
    29/07/2003 - Scarpa D. - Agregado de la columna confsuma   
    25/08/2003 - Scarpa D. - Agregado de la columna confval2		
	03/10/2003 - Scarpa D. - Cambio en el titulo de las columnas
-----------------------------------------------------------------------------
-->

<% 
Dim l_rs
Dim l_sql

Dim l_repnro
Dim l_confnrocol
Dim l_confetiq
Dim l_conftipo
dim	l_confval
dim	l_confval2
dim	l_confaccion
	
dim l_filtro
dim l_orden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
		l_orden = " ORDER BY confnrocol"
end if

l_repnro = request("repnro")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/trivoliSwimming/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n de Reportes - Ticket</title>
</head>
<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro,repnro,conftipo,confetiq,confval,confval2,confacc)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 document.datos.repnro.value = repnro;
 document.datos.conftipo.value = conftipo;
 document.datos.confetiq.value = confetiq;
 document.datos.confval.value = confval;
 document.datos.confval2.value = confval2;
 document.datos.confaccion.value = confacc; 
 
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}

</script>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th nowrap>Nro Columna</th>
        <th nowrap>Tipo</th>
        <th nowrap>Etiqueta</th>
        <th nowrap>V. Num.</th>
        <th nowrap>V. AlfaNum.</th>		
        <th nowrap>Acci&oacute;n</th>
		
    </tr>
<%


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT repnro, "  
l_sql = l_sql & " confnrocol,  "
l_sql = l_sql & " conftipo,   "
l_sql = l_sql & " confetiq,   "
l_sql = l_sql & " confval,   "
l_sql = l_sql & " confval2,   "
l_sql = l_sql & " confaccion "
l_sql = l_sql & " FROM  confrep"
l_sql = l_sql & " WHERE confrep.repnro = " & l_repnro
		
if l_filtro <> "" then
	 l_sql = l_sql & " WHERE " & l_filtro 
end if
	
l_sql = l_sql & " " & l_orden	

rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="5">No hay datos</td>
</tr>
<%else
	do until l_rs.eof
	%>
	<tr onclick="Javascript:Seleccionar(this,<%=l_rs("confnrocol")%>,<%=l_rs("repnro")%>,'<%=l_rs("conftipo")%>','<%= Server.URLEncode( l_rs("confetiq") )%>',<%=l_rs("confval")%>,'<%=l_rs("confval2")%>','<%=l_rs("confaccion")%>')">
		<td width="10%"><%=l_rs("confnrocol")%></td>
		<td width="20%"><%=l_rs("conftipo")%> </td>
		<td width="20%"><%=l_rs("confetiq")%> </td>
		<td width="20%"><%=l_rs("confval")%> </td>
		<td width="20%"><%=l_rs("confval2")%> </td>		
		<% if isNull(l_rs("confaccion")) then
 	          response.write "<td width=""10%""></td>"		
		   else		
			   if CStr(l_rs("confaccion")) = "sumar" then
			      response.write "<td width=""10%"">Sumar</td>"
			   else
			      if CStr(l_rs("confaccion")) = "restar" then
	  		         response.write "<td width=""10%"">Restar</td>"
				  else
				     response.write "<td width=""10%""> - </td>"
				  end if
			   end if
		   end if
		%>
	</tr>
	<%l_rs.MoveNext
	loop
end if ' del if l_rs.eof
l_rs.Close
cn.Close	
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0" >
<input type="Hidden" name="repnro" value="0" >
<input type="Hidden" name="conftipo" value="0" >
<input type="Hidden" name="confetiq" value="0" >
<input type="Hidden" name="confval" value="0" >
<input type="Hidden" name="confval2" value="0" >
<input type="Hidden" name="confaccion" value="0" >

<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value='<%= l_filtro %>'>
</form>

</body>
</html>
