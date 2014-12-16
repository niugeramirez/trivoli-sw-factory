<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% Server.ScriptTimeout = 720 %>
<%
'---------------------------------------------------------------------------------
'Archivo	: equipo_eva_01.asp
'Descripción: browse de empleados del Proyecto
'Autor		: CCRossi
'Fecha		: 16-12-2004
'Modificado	:  02-04-2005 - LAmadio - funcione con VB - y generac form eva.
'---------------------------------------------------------------------------------------
on error goto 0

'Variables base de datos
 Dim l_rs
 Dim l_rs1 
 Dim l_sql

'Variables filtro y orden
 dim l_filtro
 dim l_orden
 dim l_listempleados
 Dim l_evaproynro

 Dim l_nombre
 Dim l_esta ' para verificar si el empleado ya esta relacionado o no
 Dim l_evaluaciongenerada ' para verificar si el empleado tiene eval generada
 	' vERRRRRRRR
'  Dim l_sinasignar ' para verificar si el empleado tiene todos los evaluadores asignados
 
'Tomar parametros
 l_filtro = request("filtro")
 l_orden  = request("orden")
 l_listempleados = request("listempleados")
 l_evaproynro = request("evaproynro")
 
 'response.write l_listempleados & " - "
'response.write l_evaproynro 

'Body 
 if l_orden = "" then
	l_orden = " ORDER BY empleg"
 end if
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title>Empleados del Proyecto -  Gesti&oacute;n de Desempeño - RHPro &reg;</title>
</head>
<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,codigo){
 if (jsSelRow != null)
    Deseleccionar(jsSelRow);


 document.datos.cabnro.value = codigo;
 fila.className = "SelectedRow";
 fila.focus();
 jsSelRow		= fila;
 
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table id="tabla" name="tabla">
    <tr>
        <th nowrap>Empleado</th>
        <th nowrap>Apellido y Nombre</th>
        <th nowrap>En el Equipo</th>
        <th nowrap>Evaluaci&oacute;n Generada</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT ternro, empleg, terape,terape2, ternom, ternom2"
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " WHERE "
if trim(l_listempleados)="" then
l_sql = l_sql & " (EXISTS (SELECT * FROM evaproyemp WHERE evaproyemp.ternro = empleado.ternro  "
l_sql = l_sql & "				 AND evaproynro=  " & l_evaproynro & "))"
else
l_sql = l_sql & " ternro IN (" & l_listempleados & ")"
end if
if l_filtro <> "" then
	 l_sql = l_sql & " AND " & l_filtro 
end if
l_sql = l_sql & " " & l_orden	
'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td  colspan="4">No hay empleados relacionados al Evento.</td>
</tr>
<%else
	do until l_rs.eof
		l_nombre = l_rs("terape")
		if l_rs("terape2") <>"" then
		l_nombre = l_nombre & " " &l_rs("terape2") 
		end if
		if l_rs("ternom") <>"" or l_rs("ternom2") <>"" then
		l_nombre = l_nombre & ", " 
		end if
		l_nombre = l_nombre & l_rs("ternom") 
		l_nombre = l_nombre & " " &l_rs("ternom2") 
		
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT * FROM empleado "
		l_sql = l_sql & " WHERE "
		l_sql = l_sql & " (EXISTS (SELECT * FROM evaproyemp WHERE evaproyemp.ternro = empleado.ternro  "
		l_sql = l_sql & "				 AND evaproynro=  " & l_evaproynro & "))"
		l_sql = l_sql & " AND empleado.ternro = " & l_rs("ternro")
		rsOpen l_rs1, cn, l_sql, 0 
		l_esta=0
		if not l_rs1.EOF then
			l_esta=-1
		end if
		l_rs1.close	
		set l_rs1=nothing
		
		' l_sinasignar = 0 - para relac_Empl_
		l_evaluaciongenerada = 0
		if l_esta=-1 then
			Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT * FROM evadetevldor "
			l_sql = l_sql & " INNER JOIN evacab ON evadetevldor.evacabnro=evacab.evacabnro "
			l_sql = l_sql & "			AND evacab.empleado = "& l_rs("ternro")
			l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaevenro=evacab.evaevenro "
			l_sql = l_sql & "			AND evaevento.evaproynro=  " & l_evaproynro 
			rsOpen l_rs1, cn, l_sql, 0 
			if not l_rs1.EOF then
				l_evaluaciongenerada =-1
			end if
			l_rs1.close	
			set l_rs1=nothing
		else
			l_evaluaciongenerada =0
		end if
		%>
	<tr onclick="Javascript:Seleccionar(this,<%=l_rs("ternro")%>)">
		<td nowrap><%=l_rs("empleg")%> </td>
		<td nowrap><%=l_nombre%> </td>
		<td nowrap align=center><%if l_esta=-1 then%>SI<%else%>NO<%end if%></td>
		<td nowrap align=center><%if l_evaluaciongenerada=-1 then%>SI<%else%>NO<%end if%></td>
	</tr>
	<%l_rs.MoveNext
	loop
end if ' del if l_rs.eof
l_rs.Close
set l_rs = nothing

cn.Close	
set cn = nothing
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0" >
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
