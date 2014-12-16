<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'========================================================================================
'Archivo	: cambio_estado_evldrnro_eva_00.asp
'Descripción: cambiar marcas evadetevldor
'Autor		: CCRossi
'Fecha		: 11-05-2004
'========================================================================================

'parametro 
Dim l_evldrnro

'Datos del formulario
Dim l_ingreso
Dim l_habilitado
Dim l_evldorcargada
Dim l_fechaing
Dim l_fechahab
Dim l_fechacar
Dim l_horaing
Dim l_horahab
Dim l_horacar

'locales
dim l_evacabnro			
dim l_empleado			
dim l_etaprogcarga		
dim l_etaprogread		
dim l_evaseccnro

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_evldrnro = request.querystring("evldrnro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")		
l_sql = "SELECT ingreso,habilitado,evldorcargada,fechaing,fechahab,evadetevldor.fechacar,horaing,horahab,evadetevldor.horacar, "
l_sql = l_sql & " evadetevldor.evacabnro, empleado, etaprogcarga, etaprogread, evadetevldor.evaseccnro "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro= evadetevldor.evacabnro"
l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro= evadetevldor.evaseccnro"
l_sql = l_sql & " left  join evaseceta on evaseceta.evaseccnro= evadetevldor.evaseccnro "
l_sql = l_sql & "    AND  evaseceta.evatipnro= evasecc.evatipnro "
l_sql = l_sql & "    AND  evaseceta.evaetanro= evacab.evaetanro "
l_sql  = l_sql  & " WHERE evldrnro = " & l_evldrnro
rsOpen l_rs, cn, l_sql, 0 
if not  l_rs.eof then
	l_evacabnro			= l_rs("evacabnro")
	l_empleado			= l_rs("empleado")
	l_evaseccnro		= l_rs("evaseccnro")
	l_etaprogcarga		= l_rs("etaprogcarga")
	l_etaprogread		= l_rs("etaprogread")

	l_ingreso		= l_rs("ingreso")
	l_habilitado	= l_rs("habilitado")
	l_evldorcargada	= l_rs("evldorcargada")
	l_fechaing		= l_rs("fechaing")
	l_fechahab		= l_rs("fechahab")
	l_fechacar		= l_rs("fechacar")
	if not isnull(l_rs("horaing")) and trim(l_rs("horaing"))<>"" then
		l_horaing = left(l_rs("horaing"),2)& ":" & right(l_rs("horaing"),2)
	end if
	if not isnull(l_rs("horahab")) and trim(l_rs("horahab"))<>"" then
		l_horahab = left(l_rs("horahab"),2)& ":" & right(l_rs("horahab"),2)
	end if
	if not isnull(l_rs("horacar")) and trim(l_rs("horacar"))<>"" then
		l_horacar = left(l_rs("horacar"),2)& ":" & right(l_rs("horacar"),2)
	end if
end if
l_rs.Close
set l_rs=nothing

%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Cambio de Estados  -  RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario()
{
	var ingreso;
	var habilitado;
	var evldorcargada;
	
	if (document.datos.ingreso.checked)
		ingreso=-1;
	else
		ingreso=0;
	if (document.datos.habilitado.checked)
		habilitado=-1;
	else
		habilitado=0;
	if (document.datos.evldorcargada.checked)
		evldorcargada=-1;
	else
		evldorcargada=0;		
	
	
	abrirVentanaH('cambio_estado_evldrnro_eva_01.asp?evldrnro=<%=l_evldrnro%>&ingreso='+ingreso+'&habilitado='+habilitado+'&evldorcargada='+evldorcargada,'',500,500);
	opener.evaluadores.location.reload();	
	//opener.cargardatos(<%=l_evldrnro%>,<%=l_evaseccnro%>,habilitado,'<%=l_etaprogcarga%>','<%=l_etaprogread%>',0);	
	window.close();
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" >
<form name="datos" action="" method="post">

<input type="Hidden" name="evldrnro" value="<%= l_evldrnro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
  <tr height=25>
    <td class="th2">Cambio de Estados</td>
	<td class="th2" align="right">		  
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
  </tr>
  <tr>
    <td align="right"><b>Ingresó:</b></td>
    <td align="left">
		<input name="ingreso" type="Checkbox" <%if l_ingreso=-1 then%>checked <%end if%>> 
		<input class="rev" style="background : #e0e0de;" readonly type=text name="fechaing" size=8 value="<%=l_fechaing%>">		
		<input class="rev" style="background : #e0e0de;" readonly type=text name="horaing" size=4 value="<%=l_horaing%>">		
    </td>
 </tr>
  <tr>
    <td align="right"><b>Habilitado:</b></td>
    <td align="left">
		<input name="habilitado" type="Checkbox" <%if l_habilitado=-1 then%>checked <%end if%>> 
		<input class="rev" style="background : #e0e0de;" readonly type=text name="fechahab" size=8 value="<%=l_fechahab%>">		
		<input class="rev" style="background : #e0e0de;" readonly type=text name="horahab" size=4 value="<%=l_horahab%>">		
    </td>
 </tr>
  <tr>
    <td align="right"><b>Terminada:</b></td>
    <td align="left">
		<input name="evldorcargada" type="Checkbox" <%if l_evldorcargada=-1 then%>checked <%end if%>> 
		<input class="rev" style="background : #e0e0de;" readonly type=text name="fechacar" size=8 value="<%=l_fechacar%>">		
		<input class="rev" style="background : #e0e0de;" readonly type=text name="horacar" size=4 value="<%=l_horacar%>">		
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
