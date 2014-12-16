<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<%
'=====================================================================================
'Archivo	: tieneobjetivos_eva_ag_01.asp
'Descripción: 
'Autor		: CCRossi
'Fecha		: 26-11-2004
'Modificacion: 28-04-2005 CCRossi - no poder modificar su PROPIA condicion.
'=====================================================================================

'variables 
 Dim l_rs
 Dim l_rs1
 Dim l_sql
 dim l_filtro2
 dim l_nombre

 dim l_color
 dim l_yapaso  
 dim l_entro
 dim i
 dim l_tieneobj
 dim l_chequeado   
 
'parametros
 dim l_filtro
 dim l_orden
 dim l_logeadoternro
 
l_filtro = request("filtro")
l_orden  = request("orden")
'l_logeadoternro  = request("logeadoternro") ' viene el ternro del empleg de autogestion
l_logeadoternro = l_ess_ternro


if len(l_filtro) <> 0 then
	if left(l_filtro,1) <> "'" then
		l_filtro2 = "'" & l_filtro & "'"
	else
		l_filtro2 =  mid(l_filtro,2,len(request("filtro")) - 1)
	end if	
end if	

if l_orden = "" then
	l_orden = " ORDER BY empleado.terape"
end if
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title>Proceso de Gesti&oacute;n de Desempe&ntilde;o - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<style>
.autoelegida
{
	COLOR: Black;
	FONT-FAMILY: Verdana;
	FONT-SIZE: 08pt;
	BACKGROUND-COLOR: "#B0E0E6";
	padding : 2;
	padding-left : 5;
}
.autoNOelegida
{
	COLOR: Black;
	FONT-FAMILY: Verdana;
	FONT-SIZE: 08pt;
	BACKGROUND-COLOR: "#fffaf2";
	padding : 2;
	padding-left : 5;
}
</style>
<script>
var jsSelRow = null;
var color = null;

function Deseleccionar(fila)
{
 if (color!==1)
	fila.className = "MouseOutRow";
 else
	fila.className = "autoNOelegida";
	
}

function Seleccionar(fila,cabnro,evaevenro,evatevnro,empleg)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 document.datos.evaevenro.value = evaevenro;
 document.datos.empleg.value = empleg;
 if (evatevnro!==1)
 {
	fila.className = "SelectedRow";
	color=2
 }	
 else
 {
 	fila.className = "autoelegida";
 	color=1
 }	
 jsSelRow		= fila;

}

</script>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th><%if ccodelco=-1 then%>Supervisados<%else%>Empleado a Evaluar<%end if%></th>
        <th><%if ccodelco=-1 then%>Tiene Compromisos<%else%>Tiene Objetivos<%end if%></th>
    </tr>
    <form name="datos" method="post">
<%


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT  "
l_sql = l_sql & " empleado.ternro,empleado.empleg,empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2 "  
l_sql = l_sql & " FROM evacab "
l_sql = l_sql & " INNER JOIN evaevento    ON evaevento.evaevenro = evacab.evaevenro "
l_sql = l_sql & " INNER JOIN evatipoeva   ON evatipoeva.evatipnro = evaevento.evatipnro "
l_sql = l_sql & " INNER JOIN empleado   ON empleado.ternro = evacab.empleado "
l_sql = l_sql & " WHERE evacab.empleado <> " & l_logeadoternro
l_sql = l_sql & " AND  EXISTS (SELECT * FROM evadetevldor"
l_sql = l_sql & "	  WHERE  evadetevldor.evacabnro = evacab.evacabnro" 
l_sql = l_sql & "		AND  evadetevldor.evaluador = " & l_logeadoternro 
l_sql = l_sql & "   	AND (evadetevldor.habilitado = -1 "
l_sql = l_sql & "   	OR   evadetevldor.evldorcargada = -1 "
IF cejemplo=-1 then ' es ABN
l_sql = l_sql & "   	OR (evadetevldor.evatevnro  <> " & cautoevaluador
l_sql = l_sql & "      AND evadetevldor.evatevnro  <> " & cevaluador &")"
end if
l_sql = l_sql & "   	)"
l_sql = l_sql & "   	)" 
if l_filtro <> "" then
 l_sql = l_sql & " AND " & l_filtro 
end if
l_sql = l_sql & " " & l_orden	
'Response.Write l_sql & "<br>"
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan=2>
	 <%if ccodelco=-1 then%>
		No hay Supervisados para el Supervisor.
	 <%else%>
		No hay Evaluados para el Evaluador o Filtro Ingresado.
	 <%end if%>
	 </td>
</tr>
<%else
	
	l_entro=0
	i=0
	do until l_rs.eof
	
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT DISTINCT  evacab.evacabnro, evacab.tieneobj"
	l_sql = l_sql & " FROM evacab "
	l_sql = l_sql & " WHERE EXISTS (SELECT * FROM evadetevldor"
	l_sql = l_sql & "	  WHERE  evadetevldor.evacabnro = evacab.evacabnro" 
	l_sql = l_sql & "		AND  evadetevldor.evaluador = " & l_logeadoternro 
	l_sql = l_sql & "   	) AND tieneobj=-1"
	l_sql = l_sql & "   	  AND empleado= " & l_rs("ternro")
	rsOpen l_rs1, cn, l_sql, 0 
	if not l_rs1.eof then
		l_tieneobj=l_rs1("tieneobj")
	else
		l_tieneobj=0
	end if
	l_rs1.Close
	set l_rs1=nothing
	
		l_yapaso=0
		l_entro = -1
		l_nombre = l_rs("terape")
		
		if trim(l_rs("terape2"))<>"" then
			l_nombre = l_nombre & " " & trim(l_rs("terape2"))
		end if	
		if trim(l_rs("ternom"))<>"" or trim(l_rs("ternom2"))<>"" then
			l_nombre = l_nombre & ","
		end if	
		if trim(l_rs("ternom"))<>"" then
			l_nombre = l_nombre & " " & trim(l_rs("ternom"))
		end if	
		if trim(l_rs("ternom2"))<>"" then
			l_nombre = l_nombre & " " & trim(l_rs("ternom2"))
		end if	
		
		if l_tieneobj=-1 then
			l_chequeado = "CHECKED"
		else	
			l_chequeado = ""
		end if	
		
		i=i+1
		%>
		<tr>
			<td nowrap><%=l_rs("ternro")%>--<%=l_nombre%></td>
		    <td align=center><input tieneobj="<%=l_rs("ternro")%>" name="cheq<%= i%>" type="Checkbox" <%= l_chequeado%> > 
		       </td>
		</tr>
		<%
	l_rs.MoveNext
	loop
	
	if l_entro=0 then%>
	<tr>
	 <td colspan="3">
	 <%if ccodelco=-1 then%>
		No hay Supervisados para el Supervisor.
	 <%else%>
		No hay Evaluados para el Evaluador o Filtro Ingresado.
	 <%end if%>
	 </td>
	</tr>
	<%end if
	
end if ' del if l_rs.eof
l_rs.Close
cn.Close	
%>
</table>

</form>

</body>
</html>
