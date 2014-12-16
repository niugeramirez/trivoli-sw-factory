<% option explicit %>
<%
'Archivo	: firmas_consulta_00.asp
'Descripción: Consulta de Firmas
'Autor		: CCRossi
'Fecha		: 03-02-2004
'Modificacion: 
'------------------------------------------------------------------------------------

' Variables
' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

' Filtro
  l_etiquetas = "Fecha:;Hora:;Autorizado Por:;Tipo;Fin de Firma:"
  l_Campos    = "cysfirmas.cysfirfecaut;cysfirmas.cysfirmhora;cysfirmas.cysfirautoriza;cystipo.cystipnombre;cysfirmas.cysfirfin"
  l_Tipos     = "F;T;T;T;T"

' Orden
  l_Orden     = "Fecha:;Hora:;Autorizado Por:;Tipo;Fin de Firma:"
  l_CamposOr  = "cysfirmas.cysfirfecaut;cysfirmas.cysfirmhora;cysfirmas.cysfirautoriza;cystipo.cystipnombre;cysfirmas.cysfirfin"

Dim l_cysfircodext
l_cysfircodext = Request.QueryString("cysfircodext")
Dim l_cystipnro
l_cystipnro = Request.QueryString("cystipnro")
%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Consulta de Firmas&nbsp; - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>
function filtro(pag)
{
  abrirVentana('filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag)
{
  abrirVentana('orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function llamadaexcel(){
    abrirVentana('firmas_consulta_02.asp?filtro=' + escape(document.ifrm.datos.filtro.value) + '&orden=' + escape(document.ifrm.datos.orden.value)+ "&"+param() ,'execl',250,150);
}

function param(){
	chequear= "cysfircodext=<%= l_cysfircodext %>&cystipnro=<%=l_cystipnro%>";
	return chequear;
}

</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
<tr style="border-color :CadetBlue;">
<td align="left" class="barra">Firmas</td>
<td class="barra" align="right">
<a class=sidebtnSHW href="Javascript:orden('firmas_consulta_01.asp');">Orden</a>
<a class=sidebtnSHW href="Javascript:filtro('firmas_consulta_01.asp');">Filtro</a>
&nbsp;
<a class=sidebtnSHW href="Javascript:llamadaexcel();">Salida Excel</a>
&nbsp;
<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
</td>
</tr>
<tr valign="top" height="100%">
   <td style="" colspan="2">
   <iframe name="ifrm" src="firmas_consulta_01.asp?cysfircodext=<%=l_cysfircodext%>&cystipnro=<%=l_cystipnro%>" width="100%" height="100%"></iframe>
   </td>
</tr>
<tr>
    <td align="right" class="th2" colspan="2">
		<a class=sidebtnABM href="Javascript:window.close()">Aceptar</a>
	</td>
</tr>
</table>
</body>
</html>