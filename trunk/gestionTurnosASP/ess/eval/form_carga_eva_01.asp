<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'---------------------------------------------------------------------------------
'Archivo	: form_carga_eva_01
'Descripción: browse de secciones en form de Cargar
'Autor		: ?
'Fecha		: ?
'Modificado	: 13-07-2004 CCRossi. acento en titulo...
'Modificado	: 22-10-2004 CCRossi. Cambiar "titulo" por "Secciones"
'Modificacion: el manejo de logeadoempleg...
'Modificado	: 12-11-2004 CCRossi. Mostrar columna de seccion obligatoria
'Modificado	: 03-02-2005 CCRossi. Cambiar Rotulo de columna para Codelco
'---------------------------------------------------------------------------------------
on error goto 0

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_empleado
Dim l_evaevenro
Dim l_revisor
Dim l_obligatoria
Dim l_tieneobj

l_empleado		= request.querystring("ternro")
l_evaevenro		= request.querystring("evaevenro")
l_revisor		= request.querystring("revisor")
l_logeadoempleg = request.querystring("logeadoempleg")

'response.write "<script>alert('form 01 : "&l_logeadoempleg&"');</script>"	

dim l_tipsecread
dim l_tipsecprog
Dim objOpenFile, objFSO, strPath

l_orden = " ORDER BY orden "

dim l_letra
dim l_pantalla
l_pantalla = request("pantalla")
if trim(l_pantalla) = "1024" then
	l_letra="style=font-size:8pt font-type:tahoma"
else	
	l_letra="style=font-size:7pt font-type:arial"
end if

dim l_logeadoempleg
l_logeadoempleg = request.querystring("logeadoempleg")

l_tieneobj=-1
if cejemplo=-1 then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT tieneobj FROM evacab WHERE evacab.empleado = " & l_empleado & " and evacab.evaevenro =" & l_evaevenro
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_tieneobj=l_rs("tieneobj")
	end if
	l_rs.close
	set l_rs=nothing	
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Secciones del Formulario - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro,pag,evacabnro,pagread)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };
//alert(cabnro);
document.datos.cabnro.value = cabnro;
parent.evaluadores.location="form_carga_eva_02.asp?ternro=<%= l_empleado %>&evaseccnro="+cabnro+"&evacabnro="+evacabnro+"&revisor="+document.datos.revisor.value+"&pantalla=<%=l_pantalla%>&logeadoempleg=<%=l_logeadoempleg%>";
parent.cargaseccion= pag;
parent.cargaseccionread= pagread;

 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th><font <%=l_letra%>>Orden</th>
        <th><font <%=l_letra%>><%if ccodelco=-1 then%>Etapas<%else%>Secciones<%end if%></th>
        <th><font <%=l_letra%>>Obligatoria</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evasecc.evaseccnro, evasecc.orden, evasecc.titulo, evatiposecc.tipsecprog, evatiposecc.tipsecread, evacab.evacabnro, evasecc.evaoblig FROM evacab inner join evadet on evacab.evacabnro= evadet.evacabnro "
l_sql = l_sql & " inner join evasecc on evadet.evaseccnro= evasecc.evaseccnro inner join evatiposecc on evatiposecc.tipsecnro= evasecc.tipsecnro "
if l_tieneobj=0 then
l_sql = l_sql & " and evatiposecc.tipsecobj <> - 1" 
end if
l_sql = l_sql & " WHERE evacab.empleado = " & l_empleado & " and evacab.evaevenro =" & l_evaevenro
if l_filtro <> "" then
  l_sql = l_sql & " and " & l_filtro & " "
end if
l_sql = l_sql & l_orden

rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
	if trim(l_rs("tipsecprog"))="" or isnull(l_rs("tipsecprog")) then
		l_tipsecprog = ""
	else
		l_tipsecprog = l_rs("tipsecprog")
		strPath = Server.MapPath(l_tipsecprog)
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If not objFSO.FileExists(strPath) Then
		   l_tipsecprog="*"
		End If
	end if

	if trim(l_rs("tipsecread"))="" or isnull(l_rs("tipsecread")) then
		l_tipsecread = ""
	else
		l_tipsecread = l_rs("tipsecread")
		strPath = Server.MapPath(l_tipsecread)
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If not objFSO.FileExists(strPath) Then
		   l_tipsecread="*"
		End If
	end if
	if l_rs("evaoblig")=-1 then
		l_obligatoria="SI"
	else
		l_obligatoria="--"
	end if	
	%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("evaseccnro")%>,'<%= l_tipsecprog%>',<%= l_rs("evacabnro")%>,'<%= l_tipsecread%>')">
        <td><font <%=l_letra%>><%= l_rs("orden")%></td>
        <td><font <%=l_letra%>><%= l_rs("titulo")%></td>
        <td align=center><font <%=l_letra%>><%= l_obligatoria%></td>
    </tr>
<%
	l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="pgrrest" value="">
<input type="Hidden" name="revisor" value="<%= l_revisor %>">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
