<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'========================================================================================
'Archivo	: etapa_form_masiva_eva_00.asp
'Descripción: cambiar etapa a formularios
'Autor		: CCRossi
'Fecha		: 24-05-2004
'========================================================================================

'parametro 
Dim l_evaevenro

'Datos del formulario
Dim l_evatipnro
Dim l_evaetanro

'local
Dim l_evatipdesabr

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_evaevenro = request.querystring("evaevenro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")		
l_sql = "SELECT evatipnro "
l_sql = l_sql & " FROM evaevento"
l_sql  = l_sql  & " WHERE evaevenro = " & l_evaevenro
rsOpen l_rs, cn, l_sql, 0 
if not  l_rs.eof then
	l_evatipnro	   = l_rs("evatipnro")
end if
l_rs.Close
set l_rs=nothing

%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Cambio de Etapa a Formularios-  RHPro &reg;</title>
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
<form name="datos" action="etapa_form_masiva_eva_01.asp" method="post">
<input type="Hidden" name="evaevenro" value="<%= l_evaevenro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
  <tr>
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
		</select>
	
	</td>	
</tr>
<tr>
    <td align="right"><b>Todos los eventos con <br>los  formularios seleccionados:</b></td>
	<td><input type=checkbox checked  name=todos></td>	
</tr>
  <tr>
    <td align="right"><b>Formularios:</b></td>
	<td>
		<select name=evatipnro size="10" multiple>
		<%	
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evatipnro, evatipdesabr"
			l_sql  = l_sql  & " FROM evatipoeva "
			l_sql  = l_sql  & " ORDER BY evatipdesabr"
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof		%>	
			<option <%if l_evatipnro = l_rs(0) then%> selected <%end if%>
				value='<%= l_rs("evatipnro") %>'> 
			<%= l_rs("evatipdesabr") %> (<%=l_rs("evatipnro")%>)</option>
		<%			l_rs.Movenext
			loop
			l_rs.Close %>	
			
		</select>
	
	</td>	
</tr>
<script>document.datos.evatipnro.value ='<%=l_evatipnro%>';</script>
<tr>
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
