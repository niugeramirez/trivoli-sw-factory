<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<% 
'Archivo	: pedido_emp_vac_gti_01
'Descripción: Pedidos Vacaciones
'Autor		: Scarpa D.
'Fecha		: 08/10/2004
'Modificado	: 

on error goto 0

Dim l_rs
Dim l_sql

Dim l_ternro
Dim l_vdiapednro
Dim l_vdiaspedestado	

dim l_filtro
dim l_filtro2
dim l_orden

l_ternro = l_ess_ternro

l_filtro = request("filtro")
l_orden  = request("orden")

if len(l_filtro) <> 0 then
	if left(l_filtro,1) <> "'" then
		l_filtro2 = "'" & l_filtro & "'"
	else
		l_filtro2 =  mid(l_filtro,2,len(request("filtro")) - 1)
	end if	
end if	

if l_orden = "" then
		l_orden = " ORDER BY vdiapeddesde"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title>Pedido de Vacaciones - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
</head>
<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}

</script>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Per&iacute;odo</th>
        <th>Fecha Desde</th>
        <th>Fecha Hasta</th>
        <th>Cantidad</th>
        <th>Estado</th>
        <th>D&iacute;as&nbsp;H&aacute;biles</th>
        <th>D&iacute;as&nbsp;No&nbsp;H&aacute;biles</th>		
        <th nowrap>D&iacute;as&nbsp;Feriados</th>
    </tr>
<%


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT vacdesc, vdiapednro, "  
l_sql = l_sql & " vdiapeddesde,  "
l_sql = l_sql & " vdiapedhasta,  "
l_sql = l_sql & " vdiapedcant,   "
l_sql = l_sql & " vdiaspedestado,  "
l_sql = l_sql & " vdiaspedhabiles,   "
l_sql = l_sql & " vdiaspedferiados,   "
l_sql = l_sql & " vdiaspednohabiles  "
l_sql = l_sql & " FROM vacdiasped "
l_sql = l_sql & " INNER JOIN vacacion ON vacacion.vacnro = vacdiasped.vacnro"
l_sql = l_sql & " WHERE vacdiasped.ternro  = " & l_ternro
		
if l_filtro <> "" then
	 l_sql = l_sql & " AND " & l_filtro 
end if
	
l_sql = l_sql & " " & l_orden	

rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="8">No hay datos para el empleado</td>
</tr>
<%else
	do until l_rs.eof
	
	if l_rs("vdiaspedestado")  then ' es true. INFORMIX CAMBIAR POR -1 ==========
		l_vdiaspedestado = "Aceptado"
	else
		l_vdiaspedestado = "Pendiente"
	end if	
	
	if CInt(l_rs("vdiaspedestado")) = -1 then%>
	<tr onclick="Javascript:Seleccionar(this,0)">
	<%else%>
	<tr onclick="Javascript:Seleccionar(this,<%=l_rs("vdiapednro")%>)">	
	<%end if%>
		<td nowrap><%=l_rs("vacdesc")%></td>
		<td ><%=l_rs("vdiapeddesde")%></td>
		<td ><%=l_rs("vdiapedhasta")%> </td>
		<td ><%=l_rs("vdiapedcant")%> </td>
		<td align=center><%=l_vdiaspedestado%> </td>
		<td ><%=l_rs("vdiaspedhabiles")%> </td>
		<td ><%=l_rs("vdiaspednohabiles")%> </td>		
		<td ><%=l_rs("vdiaspedferiados")%> </td>
	</tr>
	<%l_rs.MoveNext
	loop
end if ' del if l_rs.eof
l_rs.Close
cn.Close	
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="" >
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro2 %>">


</form>

</body>
</html>
