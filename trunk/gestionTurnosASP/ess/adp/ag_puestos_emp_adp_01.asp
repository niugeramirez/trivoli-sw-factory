<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-------------------------------------------------------------------------------------------------
Archivo       : ag_puestos_emp_adp_01.asp.
Descripcion   : Muestra el puesto del empleado y el link a la descripción
Fecha         : 26/07/2007.
Autor         : Gustavo Ring.
Modificacion  :
-------------------------------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_rs2
Dim l_sql
Dim l_elhoradesde
Dim l_elhorahasta
Dim leg
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_estado
Dim l_ternro
Dim l_repnro 
Dim l_sql_confrep
Dim l_fecha
 
Set l_rs  = Server.CreateObject("ADODB.RecordSet")


l_ternro = l_ess_ternro
l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY estructura.estrdabr"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Puestos - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro,ternro,modelin,estado)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 document.datos.ternro.value = ternro;
 document.datos.estado.value = estado;
 document.datos.ModEli.value = modelin;
 
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th nowrap>C&oacute;digo</th>
        <th nowrap>Descripci&oacute;n</th>
        <th nowrap>Archivo Descriptivo</th>
    </tr>
<%
leg = Session("empleg")
l_fecha = now()

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT distinct estructura.estrnro, estructura.estrdabr,puesto.puearchdes "
l_sql = l_sql & " FROM puesto"
l_sql = l_sql & " INNER JOIN estructura ON puesto.estrnro = estructura.estrnro and estructura.tenro = 4 "
l_sql = l_sql & " INNER JOIN his_estructura ON his_estructura.estrnro = estructura.estrnro "
l_sql = l_sql & " INNER JOIN empleado ON his_estructura.ternro = empleado.ternro "
l_sql = l_sql & " WHERE empleg = " &  leg
l_sql = l_sql & " AND his_estructura.htetdesde <=" & cambiafecha(l_fecha,"","")
l_sql = l_sql & " AND ((his_estructura.htethasta IS NULL) " 
l_sql = l_sql & "      OR "
l_sql = l_sql & "      (his_estructura.htethasta >=" & cambiafecha(l_fecha,"","")
l_sql = l_sql & "       )) "
	
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 

if l_rs.eof then%>
<tr>
	 <td colspan="3">El empleado no tiene asociado un Puesto.</td>
</tr>
<%
'-----------------------------------------------------------------
else
	%>
	<tr> 
		<td nowrap><%=l_rs("estrnro")%></td>
		<td nowrap><%=l_rs("estrdabr")%> </td>
		<td nowrap><% if not isnull(l_rs("puearchdes")) then %><a href="<%= l_rs("puearchdes")%>">Descripción</a><%end if%>&nbsp;</td>
	</tr>
	<%
end if ' del if l_rs.eof

set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="ternro" value="0">
<input type="Hidden" name="estado" value="0">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form> 

</body>
</html>
