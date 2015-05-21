<!DOCTYPE html>
<% 
Dim l_tipo
Dim l_texto1
Dim l_texto2
Dim l_texto3
Dim l_titulo
Dim l_error

l_tipo = request("tipo")

if l_tipo <> "pass" then
	l_texto1 = "Usuario:" 
	l_texto2 = "Contraseña:"
	'l_texto3 = "Empresa:"
	l_titulo = "Identificación del Usuario"
	l_error  = ""
else
	l_texto1 = "Contraseña:"
	l_texto2 = "Nueva:"
	l_texto3 = "Confirma:"
	l_titulo = "Usuario: " & session("loguinUser")
	l_error  = "Cambiar contraseña"
end if
 %>
<html lang="en">
	<head>
		<title></title>
		<meta charset="utf-8">
		<link rel="stylesheet" href="css/reset2.css" type="text/css" media="all">
		<link rel="stylesheet" href="css/layout2.css" type="text/css" media="all">

		<link rel="stylesheet" href="css/style4.css" type="text/css" media="all">
		<script type="text/javascript" src="js/jquery-1.6.js" ></script>
		<script type="text/javascript" src="js/cufon-yui.js"></script>
		<script type="text/javascript" src="js/cufon-replace.js"></script>  
		<script type="text/javascript" src="js/Vegur_300.font.js"></script>
		<script type="text/javascript" src="js/PT_Sans_700.font.js"></script>
		<script type="text/javascript" src="js/PT_Sans_400.font.js"></script>
		<script type="text/javascript" src="js/tms-0.3.js"></script>
		<script type="text/javascript" src="js/tms_presets.js"></script>
		<script type="text/javascript" src="js/jquery.easing.1.3.js"></script>
		<script type="text/javascript" src="js/atooltip.jquery.js"></script>
		<!--[if lt IE 9]>
		<script type="text/javascript" src="js/html5.js"></script>
		<link rel="stylesheet" href="css/ie.css" type="text/css" media="all">
		<![endif]-->
		<!--[if lt IE 7]>
			<div style=' clear: both; text-align:center; position: relative;'>
				<a href="http://windows.microsoft.com/en-US/internet-explorer/products/ie/home?ocid=ie6_countdown_bannercode"><img src="http://storage.ie6countdown.com/assets/100/images/banners/warning_bar_0000_us.jpg" border="0" height="42" width="820" alt="You are using an outdated browser. For a faster, safer browsing experience, upgrade for free today." /></a>
			</div>
		<![endif]-->
		
		
<style type="text/css">
<!--
a {
	text-decoration: none;
}
.texto{
	font : Arial;
	font-size : 17px;
	color : Navy;
	text-align : right;
	border : Black;
	border: none;
	background-color : transparent;
	border:none;
}
.blanc{
	font:16px;
	border:1px solid Navy;
	background-color : transparent;
	border:none;
/*	border:1px solid #87CEFA;*/
	border:1px solid Navy;
	width:90%;
	color:Navy;
	z-index:100;
}
.blancpass{
	font:16px;
/*	border:1px solid Blue;*/
/*	background-color : transparent;*/
	border:1px solid #87CEFA;
/*	border:none;*/
	width:118px;
	color: Navy;
	height: 22px;
}

-->
</style>

<script>
var usertmp;
var loguser;
var passuser;
var domuser;

function tip(texto){
	document.FormVar.titulo.value = texto;
}
function help(){
	abrirVentana(document.FormVar.ayuda.value,'',750,550);
}

function ok(){
	t=event.keyCode;
	event.returnValue = true;
	if (t==13){
		ok2();
	}
}

function ok2(){


	<% If l_tipo <> "pass" then %>

		document.cookie = "usr=;expires=now";
		if (document.FormVar.seguridad.value == -1){
			document.FormVar.usr.value = domuser;
		}else{
			document.FormVar.usr.value = document.FormVar.usr2.value;
		}
		if (document.FormVar.empresa.value == 'desa'){
			//document.FormVar.seg_nt.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].seg;
			//document.FormVar.seguridad.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].seg;
			//document.FormVar.base.value  = document.FormVar.basex[document.FormVar.basex.selectedIndex].bases;
		}else{
			//document.all.seg_nt.value = document.FormVar.basex.seg;
			//document.all.seguridad.value = document.FormVar.basex.seg;
			//document.all.base.value = document.FormVar.basex.bases;

		}
		if (document.FormVar.usr2.value == ""){
			alert('Debe ingresar un usuario.');
			return;
		}
		document.FormVar.action = 'shared/db/default.asp';
	<% Else  %>
		document.FormVar.action = 'seguridad/cambio_clave_seg_01.asp';
	<% End If %>
		document.FormVar.target = "_self";
		document.FormVar.method = "POST";
		//alert(document.FormVar.usr.value +"="+ document.FormVar.pass.value +"="+ document.FormVar.base.value +"="+ document.FormVar.seg_nt.value+"="+ document.FormVar.menu.value+"="+ document.FormVar.debug.value);
		document.FormVar.submit();
}

function cancel(){
	<% If l_tipo <> "pass" then %>
		window.close();
	<% Else  %>
		window.location = "lanzador2.asp";
	<% End If %>
}

function web(){
	abrirVentana(document.FormVar.inicio.value,'',750,550);
}

</script>		

		
	</head>
	<body id="page1">
	<%'= RESPONSE.END %>
		<div class="main">
		
		
		
<!--header -->
			
			
					<h1><a href="in.asp" id="logo4"></a></h1>
			
	</body>
</html>