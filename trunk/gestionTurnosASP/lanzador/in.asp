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
		<link rel="stylesheet" href="css/style3.css" type="text/css" media="all">
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
alert();

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
		document.FormVar.action = '../shared/db/default.asp';
	<% Else  %>
		document.FormVar.action = '../seguridad/cambio_clave_seg_01.asp';
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

function email(){

}
function inicio(){
	<% if l_tipo <> "pass" then %>
		document.FormVar.usr2.focus();
	<% else %>
		document.FormVar.usrpass.focus();
	<% End If %>
	document.all.ifrm2.src = "lanzador.asp?menu=html&tipo=<%= l_tipo %>";
//	window.history.lenght=0;
	document.clear();
}

function baseopt(){
	<% If l_tipo <> "pass" then %>
		if (document.all.usr2.value != usertmp){
			loguser = document.FormVar.usr2.value;
			passuser = document.FormVar.pass.value;
		}
		//if (document.FormVar.basex[document.FormVar.basex.selectedIndex].seg == -1){
			//document.FormVar.usr2.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].user;
			//document.FormVar.pass.value = "";
			//document.FormVar.usr2.disabled = true;
			//document.FormVar.pass.disabled = true;
		//}else{
			//document.FormVar.usr2.value = loguser;
			//document.FormVar.pass.value = passuser;
			//document.FormVar.usr2.disabled = false;
			//document.FormVar.pass.disabled = false;
		//}
	<% Else  %>
		document.FormVar.user.value = domuser;
	<% End If %>
	//document.FormVar.base.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].bases;
	//document.FormVar.seg_nt.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].seg;
	//document.FormVar.seguridad.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].seg;
	document.FormVar.debug.value = document.FormVar.debug.value;
}
function logout(arg){
	if (arg == "0"){
		//document.ifrm.location = "lanzador/logout.asp?arg=0";
	}else{
		//abrirVentanaH('logout.asp?arg=1','',10,10);
	}
}
</script>		

		
	</head>
	<body id="page1">
	<%'= RESPONSE.END %>
		<div class="main">
		
		
		
<!--header -->
			<header>
			
				<div class="wrapper">
			
					<h1><a href="in.asp" id="logo">Megavision.com.ar</a></h1>
					<form name="FormVar" method="post" id="search">
		<input type="hidden" name="menu" value="html">
		<input type="hidden" name="base" value=2>
		<input type="hidden" name="seg_nt" value=0>
		<input type="hidden" name="usr" value="">

		
<% If l_tipo <> "pass" then %>
		<div  style="position:absolute; top:56px; left:850px; width: 150px; height: 22px;"> 
    	    <input name="usr2" class="blanc" type="Text" size="18" border="0"  onKeyPress="ok();" style="background-color:#FFFFFF;" onblur="if(this.value=='') this.value='Usuario'" onFocus="if(this.value =='Usuario' ) this.value=''" >
		</div>
        <div style="position:absolute;top:83px;left:850px; width: 150px"> 
 	       <input name="pass" type="password" size="19" class="blanc"  onKeyPress="ok();" style="background-color:#FFFFFF;" >
		</div>
        <div style="position:absolute;top:83px;left:990px; width: 50px"> 
		   <input name="btnok" type="button" value="Ingresar" onClick="ok2();" onMouseOver="tip('Aceptar');" onMouseOut="tip('');">
		</div>		
		
	<% Else  %>
		<div style="position:absolute; top:126px; left:200px; width: 50px; height: 22px;"> 
    	    <input name="usrpass" class="blancpass" type="password" size="18" border="0"  onKeyPress="ok();" style="background-color:#FFFFFF;">
		</div>
        <div style="position:absolute;top:153px;left:200px"> 
 	       <input name="usrpassnuevo" type="password"  size="19" class="blancpass"  onKeyPress="ok();" style="background-color:#FFFFFF;">
		</div>
		<div style="position:absolute;top:180px;left:200px"> 
 	       <input name="usrconfirm" type="password" size="19" class="blancpass"  onKeyPress="ok();" style="background-color:#FFFFFF;">
		</div>
	<% End If %>		
		
		
		
		
<input type="hidden" name="inicio" value="">
		<input type="hidden" name="email2" value="">
		<input type="hidden" name="ayuda" value="">
		<input type="hidden" name="empresa" value="">
		<input type="hidden" name="seguridad" value=0>
		<input type="hidden" name="conexion" value=0>
		<input type="hidden" name="debug" value=0>
		<input type="hidden" name="tipo" value=<%= l_tipo %>>
		
									

							
					
		   
		   	
						<!--
						<fieldset>
							<div class="bg"><input class="input" type="text" value="Search"  onblur="if(this.value=='') this.value='Search'" onFocus="if(this.value =='Search' ) this.value=''" ></div>
							
						</fieldset>-->
				
					</form>
				</div>


				<div id="slider">
					<ul class="items">
						<li>
							<img src="images/img1.jpg" alt="">
							<div class="banner">
								<span class="title"><span class="color2">Megavision</span><span class="color1">Centro Privado</span><span>Oftalmologia</span></span>
								<p>Nuestro Centro se especializa en el tratamientos de patologias de retina y vitreo</p>
								<a href="#" class="button1">Mas</a>
							</div>
						</li>
						<li>
							<img src="images/img2.jpg" alt="">
							<div class="banner">
								<span class="title"><span class="color2">Consultorios</span><span class="color1"> de Ultima</span><span>Generacion</span></span>
								<p>Contamos con un Quirofano especialmente disenado para cirugia de alta complejidad.</p>
								<a href="#" class="button1">Mas</a>
							</div>
						</li>
						<!--
						<li>
							<img src="images/img3.jpg" alt="">
							<div class="banner">
								<span class="title"><span class="color2">The Best</span><span class="color1">You Can Find</span><span>On The Web</span></span>
								<p>Lorem ipsum dolor sit amet, consectetur adipisicing elit sed do eiusmod tempor.</p>
								<a href="#" class="button1">Read More</a>
							</div>
						</li> -->
					</ul>
				</div>
			</header>
<!--header end-->




<!--footer -->
			<footer>
				<ul id="icons">
					<li class="first">Seguinos en:</li>
					<li><a href="#" class="normaltip" title="Facebook"><img src="images/icon1.jpg" alt=""></a></li>
					<li><a href="#" class="normaltip" title="Twitter"><img src="images/icon2.jpg" alt=""></a></li>
					<!--<li><a href="#" class="normaltip" title="Picasa"><img src="images/icon3.jpg" alt=""></a></li>-->
					<li><a href="#" class="normaltip" title="YouTube"><img src="images/icon4.jpg" alt=""></a></li>
				</ul>
				Megavision.com.ar &copy; 2015 <br>Desarrollado por <a rel="nofollow" href="http://www.templatemonster.com/" target="_blank">Trivoli</a><br>
				<!-- {%FOOTER_LINK} -->
			</footer>
<!--footer end-->
		</div>
		<script type="text/javascript"> Cufon.now(); </script>
		<script>
			$(window).load(function(){
				$('#slider')._TMS({
					banners:true,
					waitBannerAnimation:false,
					preset:'diagonalFade',
					easing:'easeOutQuad',
					pagination:true,
					duration:400,
					slideshow:8000,
					bannerShow:function(banner){
						banner.css({marginRight:-500}).stop().animate({marginRight:0}, 600)
					},
					bannerHide:function(banner){
						banner.stop().animate({marginRight:-500}, 600)
					}
					})
			})
		</script>
	</body>
</html>