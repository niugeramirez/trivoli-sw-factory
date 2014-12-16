<% Option Explicit %>
<% 
Dim l_empleg
Dim l_terape
Dim l_ternombre
Dim l_tdnro
Dim l_tddesc
Dim l_elfechadesde
Dim l_elfechahasta
Dim l_sql
Dim l_rs
%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Registraciones - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>

function menorque(fecha1,fecha2){
	var f1= new Date(); 
	f1.setFullYear(fecha1.substr(6,4),fecha1.substr(3,2)-1,fecha1.substr(0,2));
	var segf1=Date.parse(f1); 

	var f2= new Date(); 
	f2.setFullYear(fecha2.substr(6,4),fecha2.substr(3,2)-1,fecha2.substr(0,2));
	var segf2=Date.parse(f2); 

	if ((segf1<segf2)||(fecha1==fecha2)){return true}
	else{return false}
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}


function validar(){
var height
if (document.datos.fechadesde.value == "") 
	alert("Debe ingresar la fecha desde.");
else
if (document.datos.fechahasta.value == "") 
	alert("Debe ingresar la fecha hasta.");
else 
if (validarfecha(document.datos.fechadesde) && validarfecha(document.datos.fechahasta))
{
if (!menorque(document.datos.fechadesde.value,document.datos.fechahasta.value)) 
	alert("La fecha Hasta es menor que la fecha Desde.");
else 
{
 var height=330;
 var width= 550;		
 var str = "height=" + height + ",innerHeight=" + height;
  str += ",width=" + width + ",innerWidth=" + width;
  if (window.screen) {
    var ah = screen.availHeight - 30;
    var aw = screen.availWidth - 10;

    var xc = (aw - width) / 2;
    var yc = (ah - height) / 2;

    str += ",left=" + xc + ",screenX=" + xc;
    str += ",top=" + yc + ",screenY=" + yc;
  }

	window.open("","ventana",str);
	document.datos.submit();
	window.close();
}
}
}
</script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>

<% 

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="registr_diarias_gti_06.asp" target="ventana" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="th2" align="left" colspan="2">Ingreso de Fechas</td>
	<td class="th2" colspan="2" align="right"><a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
</table>

<table cellspacing="1" cellpadding="0" border="0" width="100%">
<tr>
	<td colspan="4" height="10">
		<br>
	</td>
</tr>
<tr>
    <td align="right"><b>Desde:</b></td>
	<td>
	<input type="text" name="fechadesde" size="10" maxlength="10">
	<a href="Javascript:Ayuda_Fecha(document.datos.fechadesde)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
	</td>
    <td align="right"><b>Hasta:</b></td>
	<td>
	<input type="text" name="fechahasta" size="10" maxlength="10">
	<a href="Javascript:Ayuda_Fecha(document.datos.fechahasta)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
	</td>
</tr>
<tr>
	<td colspan="4" height="10">
		<br>
	</td>
</tr>

</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="Javascript:validar(); ">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>
</body>
</html>
