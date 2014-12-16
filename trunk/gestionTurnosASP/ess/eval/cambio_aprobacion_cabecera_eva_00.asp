<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'========================================================================================
'Archivo	: cambio_aprobacion_cabecera_eva_00.asp
'Descripción: cambiar aprobacion a evacab
'Autor		: CCRossi
'Fecha		: 11-05-2004

'            14-10-2005 - Leticia Amadio -  Adecuacion a Autogestion	
'========================================================================================

'parametro 
Dim l_evacabnro

'Datos del formulario
Dim l_cabaprobada


'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_evacabnro = request.querystring("evacabnro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")		
l_sql = "SELECT cabaprobada "
l_sql = l_sql & " FROM evacab"
l_sql  = l_sql  & " WHERE evacabnro = " & l_evacabnro
rsOpen l_rs, cn, l_sql, 0 
if not  l_rs.eof then
	l_cabaprobada	   = l_rs("cabaprobada")
end if
l_rs.Close
set l_rs=nothing

%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Cambio de Aprobaci&oacute;n <%if ccodelco=-1 then%>del Proceso de Gestión<%else%>de Evaluaci&oacute;n <%end if%>-  RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario()
{
	if (document.datos.cabaprobada.checked)
	{
		var r = showModalDialog('cabaprobada_eva_00.asp?evacabnro=<%=l_evacabnro%>&cabaprobada=-1', '','dialogWidth:20;dialogHeight:20'); 
		opener.document.all.cabaprobada.value='SI';
	}	
	else	
	{
		var r = showModalDialog('cabaprobada_eva_00.asp?evacabnro=<%=l_evacabnro%>&cabaprobada=0', '','dialogWidth:20;dialogHeight:20'); 
		opener.document.all.cabaprobada.value='NO';
	}	
	window.close();
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" >
<form name="datos" action="" method="post">

<input type="Hidden" name="evacabnro" value="<%= l_evacabnro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
  <tr height=25>
    <th class="th2">Aprobar / No aprobar</th>
	<th class="th2" align="right">		  
		&nbsp; <!-- <a class=sidebtnHLP href="Javascript:ayuda('<%'= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a> -->
	</th>
  </tr>
  <tr>
    <td colspan=2 align="right"><b><%if ccodelco=-1 then%>Proceso Aprobado<%else%>Evaluaci&oacute;n Aprobada<%end if%>:</b>&nbsp;
		<input name="cabaprobada" type="Checkbox" <%if l_cabaprobada=-1 then%>checked <%end if%>> 
		&nbsp;
    </td>
</tr>
<tr height=25>
    <td  colspan="2" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>
<%

cn.Close
set cn = nothing
%>
</body>
</html>
