<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'---------------------------------------------------------------------------------
'Archivo	: rep_formulario_eva_ag_00.asp
'Descripción: Impresion del formulario desde autogestion.
'Autor		: CCRossi
'Fecha		: 21-07-2004
'Modificado	: 
' 			13-07-2005 - CCROssi - pasar parametro ternro cuando en este caso que se invoca desde
' 			el formulario. Este paraemtro y l_quien="" significa que solo hay se listara el formulario del ternro.
'            14-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'  			11/08/2006 - LA. - Se agrego la opcion de configurac de pagina.- cambiar v_empleado por empleado
'----------------------------------------------------------------------------------

Dim l_rs
Dim l_sql

Dim l_evaevenro
Dim l_ternro
Dim l_estrnro
Dim l_titulo
dim l_logeadoempleg
dim l_logeadoternro

l_evaevenro=Request.QueryString("evaevenro")
l_ternro=Request.QueryString("ternro")
l_titulo=Request.QueryString("titulo")
l_logeadoempleg=Request.QueryString("logeadoempleg")

if trim(l_logeadoempleg)<>"" then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT empleado.ternro FROM empleado WHERE empleg=" & l_logeadoempleg
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then	
		l_logeadoternro = l_rs("ternro")
	else
		l_logeadoternro=""
	end if	
	l_rs.Close
	set l_rs=nothing
end if

%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%if ccodelco=-1 then%>Proceso de Gesti&oacute;n de Desempe&ntilde;o<%else%>Formulario de Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;<%end if%></title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script>
var titulo='';
var indice = 1;

function imprimir(){
	parent.frames.ifrm.focus();
	window.print();
}

function ConfPagina(){
    var WebBrowser = '<OBJECT ID="WebBrowser' + indice + '" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
    document.all.objetos.insertAdjacentHTML('beforeEnd', WebBrowser);

    execScript("on error resume next: WebBrowser" + indice + ".ExecWB 8, -1", "VBScript");
	indice++;
	
	document.all.objetos.innerHTML = '';
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos"  action="#" method="post">
<input type="Hidden" name="ternro" value="">
<input type="Hidden" name="listemp" value="">

<table cellspacing="0" cellpadding="0" border="0" width="100%" >
  <tr>
    <th class="th2" colspan="2"><%if ccodelco=-1 then%>Proceso de <%else%>Formulario de <%end if%>Gesti&oacute;n de Desempe&ntilde;o</th>
	<td nowrap colspan="2" align="right" class="th2" valign="middle">
		<a class=sidebtnSHW href="Javascript:ConfPagina()">Conf. P&aacute;gina</a>&nbsp;&nbsp;&nbsp;
		<a class=sidebtnSHW href="Javascript:imprimir();">Imprimir</a> &nbsp;
	</td>
  </tr>
</table>  
<table cellspacing="0" cellpadding="0" border="0" width="95%" height="100%">
<tr height="95%">
    <td align="right" colspan="3">
    <%if l_titulo="Borrador" then%>
		<iframe name="ifrm" src="rep_borrador_eva_01.asp?llamadora=Auto&evaevenro=<%=l_evaevenro%>&listternro=<%=l_ternro%>&titulo=<%=l_titulo%>&logeadoternro=<%=l_logeadoternro%>" width="100%" height="100%"></iframe> 
	<%else%>
		<iframe name="ifrm" src="rep_emp_formulario_eva_01.asp?llamadora=Auto&evaevenro=<%=l_evaevenro%>&listternro=<%=l_ternro%>&titulo=<%=l_titulo%>&logeadoternro=<%=l_logeadoternro%>&ternro=<%=l_ternro%>" width="100%" height="100%"></iframe> 
	<%end if%>
	</td>
</tr>
<tr height="50%">
    <td align="right" colspan="3" height="40">
	</td>
</tr>
</table>
</form>
<div id="objetos" name="objetos">

</div>

</body>
</html>
