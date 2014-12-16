<%Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<% 
'=====================================================================================
'Archivo	: tieneobjetivos_eva_ag_00.asp
'Descripción: 
'Autor		: CCRossi
'Fecha		: 26-11-2004
'Modificacion: 
'=====================================================================================

' Variables
' de parametros entrada
  Dim l_logeadoempleg   ' es el EMPLEG viene de autogestion
  Dim l_listainicial ' lista completa  de empleados
  
' de uso local  
  Dim l_logeadoternro   ' es el TERNRO del empleg que viene de autogestion
  
' de base de datos  
  Dim l_sql
  Dim l_rs

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

' Filtro
if ccodelco=-1 then
  l_etiquetas = "Evento:;Formulario:;Supervisados a Evaluar:"
  l_Orden     = "Evento:;Formulario:;Supervisados a Evaluar:"
else
 l_etiquetas = "Evento:;Formulario:;Empleado a Evaluar:"
  l_Orden     = "Evento:;Formulario:;Empleado a Evaluar:"
end if  
  l_Campos    = "evaevento.evaevedesabr;evatipoeva.evatipdesabr;empleado.terape"
  l_Tipos     = "T;T;T"

' Orden

  l_CamposOr  = "evaevento.evaevedesabr;evatipoeva.evatipdesabr;empleado.terape"

'l_logeadoempleg = Request.QueryString("empleg")
l_logeadoempleg = l_ess_empleg
l_listainicial = Request.QueryString("listainicial")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT ternro "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " WHERE  empleg = " & l_logeadoempleg
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_logeadoternro = l_rs("ternro")
end if
	
'response.write "<script>alert('"&l_logeadoempleg&"');</script>"	
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Proceso de Gesti&oacute;n de Desempe&ntilde;o  -  Gesti&oacute;n de Desempe&ntilde;o<%if ccodelco<>-1 then%> - RHPro &reg;<%end if%></title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

function filtro(pag)
{
  abrirVentana('filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag)
{
  abrirVentana('orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function param(){
	chequear= "logeadoternro=<%= l_logeadoternro %>";
	return chequear;
}

function pantalla(){
	document.datos.pantalla.value=screen.availWidth;
}

function Devolver(){
	var i;
	var lista;
	
	var formElements = document.ifrelem.datos.elements;
	i=0;
	lista='';
	var cent;
	
	while (i<formElements.length)
	{
		cent = formElements[i].tieneobj;
		if (formElements[i].checked)
		{
			if (lista=='') 
				lista = lista + cent;
			else	
				lista = lista + ',' + cent;
		}	
		i = i + 1;
	}
	
	//alert(lista);
	
	var r = showModalDialog('tieneobjetivos_eva_ag_02.asp?evaluador=<%=l_logeadoternro%>&listainicial=<%=l_listainicial%>&lista='+lista + '&centnro=', '','dialogWidth:20;dialogHeight:10'); 
	opener.ifrm.location.reload();
    window.close();

}
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="pantalla();">
<form name=datos>
	<input type=hidden name=pantalla>
</form>
<table border="0" cellpadding="0" cellspacing="0" height="100%">
<tr style="border-color :CadetBlue;" height="5%">
	<td align="left" class="th2">
	<%if ccodelco=-1 then%>
		Definir Supervisados Con Compromisos Predefinidos
	<%else%> 
		Definir Empleados Con Objetivos Predefinidos
	<%end if%> 
	</td>
	<td align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Devolver();">Guardar</a> &nbsp;
	</td>
</tr>
<tr valign="top" height="95%">
   <td colspan="2" style="">
   <iframe name="ifrelem" src="tieneobjetivos_eva_ag_01.asp?empleg=<%=l_logeadoempleg%>" width="100%" height="100%"></iframe> 
   </td>
</tr>

</table>
</body>
</html>
