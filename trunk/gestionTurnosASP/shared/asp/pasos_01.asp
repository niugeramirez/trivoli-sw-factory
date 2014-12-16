<%Option Explicit %>

<% 

' Variables
' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

' Filtro
  l_etiquetas = "Código:;Descripción:;Cód. Externo;Cant. Personas"
  l_Campos    = "estructura.estrnro;estructura.estrdabr;estructura.estrcodext;puesto.puecantpers"
  l_Tipos     = "N;T;T;N"

' Orden
  l_Orden     = "Código:;Descripción:;Cód. Externo;Cant. Personas"
  l_CamposOr  = "estructura.estrnro;estructura.estrdabr;estructura.estrcodext;puesto.puecantpers"

' ADO
  dim l_rs
  dim l_rs1
  dim l_sql

  Dim l_tenro
  Dim l_titulo
  Dim l_codigo
  
l_codigo = request("codigo")
'l_tenro = Request.QueryString("tenro")
l_titulo = Request.QueryString("titulo")

'l_tenro = Session("l_tenro")
'if l_tenro = "" then
'	l_tenro = 4
'end if

l_tenro = 4
%>

<html>
<head>
<link href="/turnos/shared/css/tablesraul.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>  buques - Oleaginosa Moreno Hnos. S.A.</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>

<script>
function filtro(pag)
{
  abrirVentana('../shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag)
{
  abrirVentana('../shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function param(){
	chequear= "tenro=<%= l_tenro %>&asistente=1&codigo=<%= l_codigo %>";
	return chequear;
}
function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("puestos_adp_excel.asp?tenro=<%=l_tenro%>&orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',350,250);
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name=datos>
	<input type="hidden" name="tenro">
	<input type="hidden" name="seleccion">
</form>
<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%" >
<tr style="border-color :CadetBlue;">
	<td align="left" class="barra" background="//../shared/images/gen_rep/boton_02.gif"><%= l_titulo%>
		<% 'call MostrarBoton ("opcionbtn", "Javascript:abrirVentana('puestos_adp_02.asp?Tipo=A&tenro=" & l_tenro & "','',550,520);","Alta")%>
		<% 'call MostrarBoton ("opcionbtn", "Javascript:eliminarRegistro(document.ifrm,'puestos_adp_04.asp?estrnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
		<% 'call MostrarBoton ("opcionbtn", "Javascript:abrirVentanaVerif('puestos_adp_02.asp?Tipo=M&estrnro=' + document.ifrm.datos.cabnro.value,'',550,520);","Modifica")%>
	</td>
	<td  align="center" class="barra" background="//../shared/images/gen_rep/boton_02.gif">

		<%'call MostrarBoton ("opcionbtn", "Javascript:EjemploDeInput(),'',500,250);","Empleados")%>
		<% 'call MostrarBoton ("opcionbtn", "Javascript:abrirVentana('sugerencias_gen_02.asp','',550,300);","Agregar Sugerencia")%>
		<%'call MostrarBoton ("opcionbtn", "Javascript:orden('../../adp/asis_puestos_adp_01.asp');","Orden")%>
		<%'call MostrarBoton ("opcionbtn", "Javascript:filtro('../../adp/asis_puestos_adp_01.asp')","Filtro")%>
	</td>
</tr>
<tr valign="top" >
   <td colspan="2" style="" height="100%" width="100%" >
   <iframe scrolling="no" name="ifrm" src="pasos_011.asp" width="100%" height="100%"></iframe>
   </td>
</tr>
</table>

</body>
</html>
