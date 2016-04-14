<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--
-->
<% 
 Dim l_wiznro
 Dim l_sql
 Dim l_rs
 Dim l_wizdesabr
 Dim l_programa

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

 Dim l_empresa
  Set l_rs  = Server.CreateObject("ADODB.RecordSet")

  
' Filtro
  l_etiquetas = "C&oacute;digo:;Descripción:"
  l_Campos    = "acunro;acudesabr"
  l_Tipos     = "N;T"

' Orden
  l_Orden     = "C&oacute;digo:;Descripción:"
  l_CamposOr  = "acunro;acudesabr"

  l_wiznro = request("wiznro")
  
  l_sql = "SELECT wizdesabr " &_
          "FROM rh_wizzard " &_
		  "WHERE wiznro = " & l_wiznro
  rsOpen l_rs, cn, l_sql, 0 
  l_wizdesabr = l_rs("wizdesabr")
  l_rs.close
  
  l_sql = "SELECT pasasp " &_
          "FROM pasos " &_
		  "WHERE wiznro = " & l_wiznro & " " &_
		  "ORDER BY pasorden"
  rsOpen l_rs, cn, l_sql, 0 
  l_programa = l_rs("pasasp")
  l_rs.close
  
  l_sql = "SELECT * FROM empresa WHERE id = " & Session("empnro")   
  rsOpen l_rs, cn, l_sql, 0 
  l_empresa = l_rs("nombre")
  l_rs.close
  
  
  
  
%>
<html>
<head>
<!--<link href="../css/tablesnuevo.css" rel="StyleSheet" type="text/css">-->

		<meta charset="utf-8">
		<link rel="stylesheet" href="css/reset.css" type="text/css" media="all">
		<link rel="stylesheet" href="css/layout.css" type="text/css" media="all">
		<link rel="stylesheet" href="css/style2.css" type="text/css" media="all">

<title><%'= Session("Titulo")%> Gesti&oacute;n Ventas</title>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>



<script>
function ActPasos(codigo, clabel, nombre){
	//alert(codigo);
	document.pasos.location = "/trivoliSwimming/shared/asp/asistente_01_RH.asp?wiznro=<%=l_wiznro%>&codigo="+codigo+"&label="+clabel+"&nombre="+nombre;
}

function RefrescarPasos() {
    //alert("Refrescandoup!!!");
    document.pasos.location.reload();
}

function Abrir(pagina, codigo, pasonro) {

  if (pagina.indexOf('?') < 0){
     document.ifrm.location = pagina+"?codigo="+codigo+"&pasnro="+pasonro;
  }else{
    document.ifrm.location = pagina+"&codigo="+codigo+"&pasnro="+pasonro;
  }    
}

function orden(pag)
{
  abrirVentana('orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag)
{
  abrirVentana('filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("acumulador_liq_excel.asp?orden="+document.ifrm.datos.orden.value+"&filtro="+escape(document.ifrm.datos.filtro.value),'execl',250,150);
}

function logout(arg){
	abrirVentanaH('logout.asp?arg=1','',10,10);
}

</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos" method="post">
<input type="hidden" name="pasonro" value="">
<input type="hidden" name="menunro" value="">
<input type="hidden" name="menunroant" value="">
</form>

<body id="page1">
		<div class="body1">
			<div class="main">
<!-- header -->
				<header>
					<h1><a href="index.html" id="logo33no"></a></h1>
					<div class="wrapper">
						<ul id="icons">
							<li><img src="images/man_48.png" alt=""></li>
							<li>Usuario: &nbsp;<%=  Session("loguinUser") %>&nbsp;(&nbsp;<%=  l_empresa %> &nbsp;)&nbsp;</li>
							<li> | </li>
							<li><a href="Javascript:logout(1);window.location= '../../index.asp';">Salir</a> </li>
							<!--
							<li><a href="#" class="normaltip" title="Twitter"><img src="images/icon2.jpg" alt=""></a></li>
							<li><a href="#" class="normaltip" title="Linkedin"><img src="images/icon3.jpg" alt=""></a></li>
							-->
						</ul>
					</div>
					
				</header>
<!-- / header -->
			</div>
		</div>

<%'= RESPONSE.END %>

<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">

	<tr>
	
		<td  bgcolor="#F1F2F2"  colspan="2">
		&nbsp;
				   
		</td>
				<!--
		
    	<td align="left" class="barra" colspan="1" height="40" style="#55739C; border-bottom: 1px solid White;"  nowrap>
		    <!--<img style="filter:Shadow(Color=White,Direction=120);" src="../images/gen_rep/tablero7.jpg"><SPAN 
			    STYLE="position: absolute; top:12px; left:260px; font-size: 18px; color: White; font-family: Arial, Helvetica, sans-serif;">
				 &nbsp;</SPAN>
				   
		</td>-->
		<!--
		<td class="barra" style="#55739C; border-bottom: 1px solid White;">
			<div align="right" >
				<table  border="0" cellspacing="0"  cellpadding="0" bgcolor="navy" width="0" height="0" >
					<tr valign="middle" >
						<td class="barra" align="right"><!-- <img src="../images/gen_rep/boton_0581.gif"> --> </td>
						<!--<td class="barra" background="../shared/images/gen_rep/boton_05.gif" align="center" width="0"><a class="opcionbtn" href="Javascript:ayuda('<%'= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a></td>
						<!--<td class="barra"><img src="../images/logosamsa.jpg"></td>
						<td class="barra"><img src="../images/logosamsa.jpg"></td>
						<td class="barra">&nbsp;&nbsp;&nbsp;</td>
					</tr>
				</table>
			</div>
		</td>-->
		
	</tr>
	<tr>
		<td nowrap align="center" width="160">
			<iframe name="pasos" src="menu1.asp" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" align="left"></iframe> 
		</td>
    	<td nowrap align="right" width="100%" height="100%">
			<!--<iframe name="ifrm" src="../../config/ventas_con_00.asp" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no"></iframe> -->
			<iframe name="ifrm" src="../../reportes/rep_estadisticas_rep_02.asp?qfechadesde=01/01/2010&qfechahasta=01/01/2020&idrecursoreservable=1" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no"></iframe>			

		</td>
	</tr>
</table>
</body>
</html>
