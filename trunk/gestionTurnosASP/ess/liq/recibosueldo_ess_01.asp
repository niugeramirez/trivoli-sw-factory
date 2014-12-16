<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!-- 
'================================================================================
'Archivo		: recibosueldo_ess_01.asp
'Descripción	: Recibos de Suedos del Empleado
'Autor			: GdeCos
'Fecha			: 14-04-2005
'Modificado		:
'================================================================================
 -->
<%
on error goto 0

Dim l_rs
Dim l_sql
Dim l_empleg

l_empleg = l_ess_empleg

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Recibos de Sueldos - Autogestion - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,recnro)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = recnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Per&iacute;odo</th>
        <th>Proceso</th>		
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

    l_sql = 		"SELECT ternro, bpronro, pliqmes, pliqanio, prodesc "
	l_sql = l_sql & "FROM rep_recibo "
	l_sql = l_sql & "WHERE rep_recibo.legajo = " & l_empleg
    l_sql = l_sql & " ORDER BY pliqanio DESC, pliqmes DESC"
	
	l_rs.MaxRecords = c_MaxRecibos

	rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="6">No existen Recibos de Sueldos.</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("bpronro")%>)">
	        <td width="30%" align="center"><%= l_rs("pliqmes")%> / <%= l_rs("pliqanio")%></td>
	        <td width="70%" nowrap><%= l_rs("prodesc")%></td>
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
</form>
</body>
</html>
