<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!-- ----------------------------------------------------------------------------
Archivo		: empleados_reportan_mss_01.asp
Descripción	: Listado de empleados que reportan al empleado logeado
Autor		: Fernando Favre
Fecha		: 30-05-2005
Modificado	:
-----------------------------------------------------------------------------  -->
<%
on error goto 0
 
 Dim l_rs
 Dim l_sql
 Dim l_empleg
 Dim l_apnom
 Dim l_ternro
 Dim l_filtro
 Dim l_orden
 
'Parámetros
 l_filtro = request("filtro")
 l_orden  = request("orden")
 
'Orden por defecto
 if l_orden = "" then
 	l_orden = " ORDER BY empleg "
 end if
 
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
 l_ternro = 0
 if Session("empleg")="" then
	l_empleg = 0
 else
	l_empleg = CInt(Session("empleg"))
 	
 	l_sql =	"SELECT ternro FROM empleado WHERE empleg = " & l_empleg
 	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
 		l_ternro = l_rs("ternro")
 	end if
 	l_rs.Close
 end if
 
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_estiloTabla%>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Empleados que reportan - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,recnro){
	if (jsSelRow != null)
		Deseleccionar(jsSelRow);

	document.datos.cabnro.value = recnro;
	fila.className  = "SelectedRow";
 	jsSelRow		= fila;
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Legajo</th>
        <th>Apellido y Nombre</th>		
    </tr>
<%
 
 l_sql = 		 "SELECT ternro, empleg, terape, ternom, terape2, ternom2 "
 l_sql = l_sql & "FROM empleado "
 l_sql = l_sql & "WHERE empreporta = " & l_ternro
 if l_filtro <> "" then
 	l_sql = l_sql & " AND " & l_filtro & " "
 end if
 l_sql = l_sql & l_orden
 
 rsOpen l_rs, cn, l_sql, 0 
 if l_rs.eof then
 	%>
	<tr>
		 <td colspan="2">No se encontraron datos</td>
	</tr>
	<%
 else
	do until l_rs.eof
		l_apnom = l_rs("terape")
		if l_rs("terape2") <> "" then
			l_apnom = l_apnom & " " & l_rs("terape2")
		end if
		l_apnom = l_apnom & ", " & l_rs("ternom")
		if l_rs("ternom2") <> "" then
			l_apnom = l_apnom & " " & l_rs("ternom2")
		end if
	%>
	    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("ternro")%>)">
	        <td width="30%" align="right"><%= l_rs("empleg")%></td>
	        <td width="70%" nowrap><%= l_apnom%></td>
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
