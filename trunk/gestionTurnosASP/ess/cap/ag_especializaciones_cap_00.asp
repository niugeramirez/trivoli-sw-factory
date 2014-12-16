<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/antigfec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: ag_especializaciones_cap_00.asp
Descripcion: especializaciones
Autor: Lisandro Moro
Fecha: 29/03/2004
-----------------------------------------------------------------------------
-->
<% 
'on error goto 0

Dim l_sql
Dim l_rs
Dim rs9
Dim l_terape 
Dim l_ternom 
Dim l_terape2 
Dim l_ternom2 
Dim l_empleg
Dim l_puenro

Dim l_empfoto
Dim rs1
Dim sql
Dim l_orinro
Dim l_estnro

Dim l_ternro

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

' Filtro
  l_etiquetas = "Cod. Elemento:; Desc. Elemento:;Cod. Nivel:;Desc. Nivel:;"
  l_Campos    = "especemp.eltananro;eltanadesabr;especemp.espnivnro;espnivdesabr;"
  l_Tipos     = "N;T;N;T;"

' Orden
  l_Orden     = "Cod. Elemento:; Desc. Elemento:;Cod. Nivel:;Desc. Nivel:;"
  l_CamposOr  = "especemp.eltananro;eltanadesabr;especemp.espnivnro;espnivdesabr;"

l_ternro = l_ess_ternro
l_empleg = l_ess_empleg

%>

<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<title> Especializaciones - Capacitación - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/menu_def.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_buscar_emp.js"></script>
<script>
function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString());
}

function emplerror(nro){
	alert('empleado error:'+nro);
}

function datos(){
var dat='';
 dat='ternro='+document.datos.ternro.value+"&empleg=<%= request.querystring("empleg")%>&empleado="+document.datos.empleado.value;
 dat= dat+ '&fechadesde='+document.datos.fecha.value+ '&fechahasta='+document.datos.fecha.value;
 return dat;
}

function ActualizarGap(){
	document.ifrm.location ="ag_especializaciones_cap_01.asp?ternro=" + document.datos.ternro.value;
}	

function llamadaexcel(){ 
	abrirVentana("ag_especializaciones_cap_excel.asp?ternro=" + document.datos.ternro.value+"&orden="+document.ifrm.datos.orden.value+"&filtro="+escape(document.ifrm.datos.filtro.value),'excel',50,50);
}

function orden(pag)
{
  abrirVentana('../shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag)
{
  abrirVentana('../shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function param(){
	chequear= "ternro=" + document.datos.ternro.value;
	return chequear;
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" >

<%
Dim l_convnro

Set l_rs = Server.CreateObject("ADODB.RecordSet")

%>
<form name="datos" action="" method="post">
<input type="hidden" name="ternro" value="<%= l_ternro %>">
<%

Dim NombrePuesto

l_sql = " SELECT estructura.estrdabr, puesto.puenro "
l_sql = l_sql & " FROM his_estructura "
l_sql = l_sql & " INNER JOIN puesto ON puesto.estrnro = his_estructura.estrnro "
l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro AND his_estructura.tenro=4 AND his_estructura.ternro=" & l_ternro & " AND his_estructura.htetdesde <= " & cambiafecha(date(),"YMD",true) & " AND ((" & cambiafecha(date(),"YMD",true) & " <= his_estructura.htethasta) OR (his_estructura.htethasta IS NULL)) "

rsOpen l_rs, cn, l_sql, 0

If Not l_rs.EOF Then
	NombrePuesto = l_rs("estrdabr")
	l_puenro     = l_rs("puenro")
else 	
	NombrePuesto = "Sin Puesto"
	l_puenro     = 0
end If

l_rs.close

dim salir 

%>

<table border="0" cellpadding="0" cellspacing="0"  height="100%">
  <tr style="border-color :CadetBlue;">
	<td colspan="2">
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
        		<th colspan="2" align="right" valign="middle">
				<a  class=sidebtnABM href="Javascript:abrirVentana('ag_especializaciones_cap_02.asp?ternro=<%= l_ternro %>&empleg=<%= request.querystring("empleg") %>&nombre=<%= l_terape & " " & l_terape2 & ", " & l_ternom & " " & l_ternom2 %>','',750,580);">Especializaciones</a>				
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;				
                <a class=sidebtnSHW href="Javascript:orden('../../cap/ag_especializaciones_cap_01.asp');">Orden</a>
		  		<a class=sidebtnSHW href="Javascript:filtro('../../cap/ag_especializaciones_cap_01.asp')">Filtro</a>
		  		&nbsp;&nbsp;														
				</th>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2">
		<table  border="0" cellpadding="0" cellspacing="0" >
			<tr>
			    <td nowrap align="right" ><b>Puesto:</b></td>
				<td><input style="background : #e0e0de;" type="text" name="convenio" size="71" maxlength="50" value='<%= NombrePuesto %>' readonly></td>
			    <td nowrap align="right">&nbsp;</td>
				<td >&nbsp;</td>
			</tr>
		</table>
	</td>
</tr>

<tr>
	<td colspan="2" height="90%">
		<table border="0" cellpadding="0" cellspacing="0"  height="100%">
	        <tr valign="top">
	        	<td align="left" style="">
    	  				<iframe  scrolling="Yes" name="ifrm" src="ag_especializaciones_cap_01.asp?ternro=<%= l_ternro %>" width="100%" height="100%"></iframe> 
     				</td>        				
      			</tr>
		</table>
	</td>
</tr>
</table>

<% 
cn.Close
set cn = Nothing
%>
</form>	
<SCRIPT SRC="/serviciolocal/shared/js/menu_op.js"></SCRIPT>
</body>
</html>
