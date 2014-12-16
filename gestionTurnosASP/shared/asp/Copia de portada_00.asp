<%Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
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
  l_etiquetas = "Apellido:;Nombre;"
  l_Campos    = "empape;empnom;"
  l_Tipos     = "T;T;"

' Orden
  l_Orden     = "Apellido:;Nombre;"
  l_CamposOr  = "empape;empnom;"

' ADO
  dim l_rs
  dim l_rs1
  dim l_sql

  Dim l_tenro
  Dim l_titulo
  Dim l_codigo
  Dim l_iduser
  Dim l_vennro
  
l_codigo = request("codigo")
'l_tenro = Request.QueryString("tenro")
l_titulo = Request.QueryString("titulo")

l_tenro = 4





%>

<html>
<head>
<link href="/turnos/shared/css/tables4.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title> aasd</title>
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
	chequear= "mes="+document.datos.mes.value+"&tenro=<%= l_tenro %>&asistente=1&codigo=<%= l_codigo %>";
	return chequear;
}
function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentanaH("cumpleanos_gen_excel.asp?mes="+ document.datos.mes.value+ "&orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',50,50);
} 	   

function Actualizar(){
	document.ifrm.location = "cumpleanos_gen_01.asp?mes=" + document.datos.mes.value;	
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name=datos>
	<input type="hidden" name="tenro">
	<input type="hidden" name="seleccion">
<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%" >
<tr style="border-color :CadetBlue;">
	<td align="center" class="barra" background="//../shared/images/gen_rep/boton_02.gif"><%= l_titulo%>
	</td>
	<td  align="right" class="barra" background="//../shared/images/gen_rep/boton_02.gif">
	</td>
</tr>
<tr valign="top" style="background-color: #000000; color: white;" >
   <td colspan="2" style="background-color: #000000; color: white;" height="100%" width="100%" >
   <iframe scrolling="no" name="ifrm" src="portada_01.asp?mes=<%=  Month(date) %>&asistente=1&tenro=<%=l_tenro%>&codigo=<%= l_codigo %>" width="100%" height="100%" ></iframe>
   </td>
</tr>
</table>
</form>
</body>
</html>
