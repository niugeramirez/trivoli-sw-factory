<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'========================================================================================
'Archivo	: etapa_cabecera_eva_00.asp
'Descripción: cambiar etapa a evacab
'Autor		: CCRossi
'Fecha		: 31-05-2004
'========================================================================================

'parametro 
Dim l_evacabnro

'Datos del formulario
Dim l_evaetanro

'local

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_evacabnro = request.querystring("evacabnro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")		
l_sql = "SELECT evaetanro "
l_sql = l_sql & " FROM evacab"
l_sql  = l_sql  & " WHERE evacabnro = " & l_evacabnro
rsOpen l_rs, cn, l_sql, 0 
if not  l_rs.eof then
	l_evaetanro	   = l_rs("evaetanro")
end if
l_rs.Close
set l_rs=nothing

%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Cambio de Etapa de Evaluaci&oacute;n -  RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario()
{
document.datos.submit();			
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" >
<form name="datos" action="etapa_cabecera_eva_01.asp" method="post">
<input type="Hidden" name="evacabnro" value="<%= l_evacabnro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
  <tr height=25>
    <td class="th2">Cambio de Etapa</td>
	<td class="th2" align="right">		  
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
  </tr>
  <tr>
    <td align="right"><b>Etapa:</b></td>
	<td>
		<select name=evaetanro size="1">
		<%	
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evaetanro, evaetadesabr"
			l_sql  = l_sql  & " FROM evaetapas "
			l_sql  = l_sql  & " ORDER BY evaetadesabr"
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof		%>	
			<option value='<%= l_rs("evaetanro") %>'> 
			<%= l_rs("evaetadesabr") %> (<%=l_rs("evaetanro")%>)</option>
		<%			l_rs.Movenext
			loop
			l_rs.Close 
			set l_rs=nothing%>
			<script>document.datos.evaetanro.value='<%=l_evaetanro%>'</script>	
		</select>
	
	</td>	
</tr>
<tr height=25>
    <td  colspan="3" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>
<%
set l_rs = nothing
'l_Cn.Close
'set l_Cn = nothing
%>
</body>
</html>
