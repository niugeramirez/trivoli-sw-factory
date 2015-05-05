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
		<link rel="stylesheet" href="css/reset.css" type="text/css" media="all">
		<link rel="stylesheet" href="css/layout.css" type="text/css" media="all">
		<link rel="stylesheet" href="css/style2.css" type="text/css" media="all">
		<script type="text/javascript" src="js/jquery-1.6.js" ></script>
		<script type="text/javascript" src="js/cufon-yui.js"></script>
		<script type="text/javascript" src="js/cufon-replace.js"></script>
  		<script type="text/javascript" src="js/Didact_Gothic_400.font.js"></script>
		<script type="text/javascript" src="js/jquery.nivo.slider.pack.js"></script>
		<script type="text/javascript" src="js/atooltip.jquery.js"></script>
		<script type="text/javascript" src="js/jquery.jqtransform.js" ></script>
		<script type="text/javascript" src="js/script.js"></script>
		<script type="text/javascript" src="js/jquery.easing.1.3.js"></script>
		<!--[if lt IE 9]>
			<script type="text/javascript" src="js/html5.js"></script>
			<style type="text/css">
				.bg{ behavior: url(js/PIE.htc); }
			</style>
		<![endif]-->
		<!--[if lt IE 7]>
			<div style=' clear: both; text-align:center; position: relative;'>
				<a href="http://www.microsoft.com/windows/internet-explorer/default.aspx?ocid=ie6_countdown_bannercode"><img src="http://www.theie6countdown.com/images/upgrade.jpg" border="0"  alt="" /></a>
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
		<div class="body1">
			<div class="main">
<!-- header -->
				<header>
					<h1><a href="index.html" id="logo3"></a></h1>
					<div class="wrapper">
					<h1><a href="index.asp" id="logo">Megavision.com.ar</a></h1>
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
					</div>
					
				</header>
<!-- / header -->
			</div>
		</div>	
	
<!-- content -->
		<div class="body2">
			<div class="main">
				<section id="content">
					<div class="wrapper">
						<article class="col1">
							<div id="slider">
								<img src="images/img1.jpg" alt="" title="<strong>Villa Neverland, 2006</strong><span>9 rooms, 3 baths, 6 beds, building size: 5000 sq. ft. &nbsp; &nbsp; &nbsp; Price: $ 600 000 &nbsp; &nbsp; &nbsp; <a href='#'>Read more</a></span>">
								<img src="images/img2.jpg" alt="" title="<strong>Villa Lipsum, 2008</strong><span>8 rooms, 4 baths, 4 beds, building size: 4500 sq. ft. &nbsp; &nbsp; &nbsp; Price: $ 500 000 &nbsp; &nbsp; &nbsp; <a href='#'>Read more</a></span>">
								<img src="images/img3.jpg" alt="" title="<strong>Villa Dolor Sid, 2007</strong><span>11 rooms, 3 baths, 5 beds, building size: 6000 sq. ft. &nbsp; &nbsp; &nbsp; Price: $ 650 000 &nbsp; &nbsp; &nbsp; <a href='#'>Read more</a></span>">
								<img src="images/img4.jpg" alt="" title="<strong>Villa Nemo Enim, 2010</strong><span>5 rooms, 2 baths, 4 beds, building size: 3000 sq. ft. &nbsp; &nbsp; &nbsp; Price: $ 400 000 &nbsp; &nbsp; &nbsp; <a href='#'>Read more</a></span>">
								<img src="images/img5.jpg" alt="" title="<strong>Villa Nam Libero, 2003</strong><span>7 rooms, 4 baths, 6 beds, building size: 7000 sq. ft. &nbsp; &nbsp; &nbsp; Price: $ 700 000 &nbsp; &nbsp; &nbsp; <a href='#'>Read more</a></span>">
							</div>
						</article>
						<article class="col2">
							<form id="form_1" method="post">
								<div class="pad1">
									<h3>Find Your Property</h3>
									<div class="row">
										City &amp; State, or Zip:<br>
										<input type="text" class="input">
									</div>
									<div class="row_select">
										Search Area:<br>
										<select><option>&nbsp;</option><option>...</option><option>...</option></select>
									</div>
									<div class="row_select">
										<div class="cols">
											Price Range:<br>
											<select><option>&nbsp;</option><option>...</option><option>...</option></select>
										</div>
										<div class="cols pad_left1">
											to:<br>
											<select><option>&nbsp;</option><option>...</option><option>...</option></select>
										</div>
									</div>
									<div class="row_select">
										<div class="cols">
											Bedroom(s):<br>
											<select><option>&nbsp;</option><option>...</option><option>...</option></select>
										</div>
										<div class="cols pad_left1">
											Bathroom(s):<br>
											<select><option>&nbsp;</option><option>...</option><option>...</option></select>
										</div>
									</div>
									<div class="row_select pad_bot1">
										<div class="cols">
											Radius:<br>
											<select><option>&nbsp;</option><option>...</option><option>...</option></select>
										</div>
										<div class="cols pad_left1">
											<a href="#" class="button" onClick="document.getElementById('form_1').submit()">Proposals</a>
										</div>
									</div>
									Know exactly what you want? <br>
									Try our <a href="#">Advanced Search</a>
								</div>
							</form>
						</article>
					</div>
					<div class="wrapper">
						<article class="col1">
							<div class="pad_left1">
								<h2 class="pad_bot1">Buyers. Sellers. Proprietors. Agents.</h2>
								<div class="wrapper">
									<article class="cols">
										<h4 class="img1">Selling</h4>
										<p class="pad_bot1"><strong class="color1">Welcome one of <br>
												<a href="http://blog.templatemonster.com/free-website-templates/" target="_blank">free website templates</a></strong></p>
										<p class="pad_bot2">
												created by TemplateMonster.com team, optimized for 1024X768 px.</p>
										<a href="#" class="button">Read more</a>
									</article>
									<article class="cols pad_left1">
										<h4 class="img2">Investing</h4>
										<p class="pad_bot1"><strong class="color1"><a href="http://blog.templatemonster.com/2011/10/10/free-website-template-slideshow-real-estate/">Free website template</a> for<br>
												Real Estate business</strong></p>
										<p class="pad_bot2">
												goes with PSD source files and without them.</p>
										<a href="#" class="button">Read more</a>
									</article>
									<article class="cols pad_left1">
										<h4 class="img3">Renting</h4>
										<p class="pad_bot1"><strong class="color1">Template has several pages</strong></p>
										<p class="pad_bot2">
												<a href="index.html" class="color2">Main</a>, <a href="Buying.html" class="color2">Buying</a>, <a href="Selling.html" class="color2">Selling</a>, <a href="Renting.html" class="color2">Renting</a>, <a href="Finance.html" class="color2">Finance</a>, <a href="Contacts.html" class="color2">Contacts</a> (note that contact us form â€“ doesnâ€™t work).</p>
										<a href="#" class="button">Read more</a>
									</article>
								</div>
							</div>
						</article>
						<article class="col2">
							<div class="pad1">
								<h3>Special Offers</h3>
								<ul class="list1">
									<li><a href="#">Home Loan Offer</a></li>
									<li><a href="#">Free Calculators</a></li>
									<li><a href="#">Free Loan Tools</a></li>
									<li><a href="#">Value Your Home</a></li>
									<li><a href="#">Recently Sold Properties</a></li>
									<li><a href="#">Suburb Statistics</a></li>
									<li><a href="#">Compare Property Prices</a></li>
								</ul>
							</div>
						</article>
					</div>
				</section>
			</div>
		</div>
		<div class="body3">
			<div class="main">
				<section id="content2">
					<div class="wrapper">
						<article class="col1">
							<div class="pad2">
								<h2>Remodeling Rooms</h2>
								<div class="wrapper pad_bot3">
									<figure class="left marg_right1"><img src="images/page1_img4.jpg" alt=""></figure>
									<p class="pad_bot1"><strong class="color2">2006, 3 baths, 6 beds, building size: 5000 sq. ft.<br>
											Price: <span class="color1">$ 600 000</span></strong></p>
									<p class="pad_bot2">
											Massa dictum ementum velitumo od consequat ante oremsumas ame consectetueraipiscing eliteli ueedlorAliquam conguen nisauris accusalla vel deinol tincidunt ligula magna semper ipsum.</p>
									<a href="#" class="button">Read more</a>
								</div>
								<div class="wrapper">
									<figure class="left marg_right1"><img src="images/page1_img5.jpg" alt=""></figure>
									<p class="pad_bot1"><strong class="color2">2006, 3 baths, 6 beds, building size: 5000 sq. ft.<br>
											Price: <span class="color1">$ 600 000</span></strong></p>
									<p class="pad_bot2">
											Massa dictum ementum velitumo od consequat ante oremsumas ame consectetueraipiscing eliteli ueedlorAliquam conguen nisauris accusalla vel deinol tincidunt ligula magna semper ipsum.</p>
									<a href="#" class="button">Read more</a>
								</div>
							</div>
						</article>
						<article class="col2">
							<div class="pad1">
								<h3>Recent News</h3>
								<div class="wrapper">
									<span class="date"><strong>28</strong><span>may</span></span>
									<p><a href="#" class="color2">Donec consequat risus</a><br>
											Hendrerit conghdim entum diam metus fringilla nisl, in porta sapien purus quis odiosem... <a href="#">more</a></p>
								</div>
								<div class="wrapper">
									<span class="date"><strong>25</strong><span>may</span></span>
									<p><a href="#" class="color2">Nullam dignissim</a><br>
											Laoreet neque, quis sollicitudin orci tempus etiam viverra leo euismod pulvinar accumsan...   <a href="#">more</a></p>
								</div>
								<div class="wrapper">
									<span class="date"><strong>22</strong><span>may</span></span>
									<p><a href="#" class="color2">Quisque nunc lorem</a><br>
											Feugiat nec sodales quis, iaculis sed libero. Cras vel nisl justo, ac posuere est nulla facilisi... <a href="#">more</a></p>
								</div>
							</div>
						</article>
					</div>
				</section>
			</div>
		</div>
<!-- / content -->
		<div class="body4">
			<div class="main">
<!-- footer -->
				<footer>
					<span class="call">Call Center: <span>1-800-567-8934</span></span>
					<a rel="nofollow" href="http://www.templatemonster.com/" target="_blank">Website template</a> designed by TemplateMonster.com<br>
					<a href="http://www.templates.com/product/3d-models/" target="_blank">3D Models</a> provided by Templates.com
					<!-- {%FOOTER_LINK} -->
				</footer>
<!-- / footer -->
			</div>
		</div>
		<script type="text/javascript"> Cufon.now(); </script>
		<script type="text/javascript">
		   $(window).load(function() {
		   $('#slider').nivoSlider({
				effect:'sliceUpDown', //Specify sets like: 'fold,fade,sliceDown, sliceDownLeft, sliceUp, sliceUpLeft, sliceUpDown, sliceUpDownLeft'
				slices:17,
				animSpeed:500,
				pauseTime:6000,
				startSlide:0, //Set starting Slide (0 index)
				directionNav:false, //Next & Prev
				directionNavHide:false, //Only show on hover
				controlNav:true, //1,2,3...
				controlNavThumbs:false, //Use thumbnails for Control Nav
				controlNavThumbsFromRel:false, //Use image rel for thumbs
				controlNavThumbsSearch: '.jpg', //Replace this with...
				controlNavThumbsReplace: '_thumb.jpg', //...this in thumb Image src
				keyboardNav:true, //Use left & right arrows
				pauseOnHover:true, //Stop animation while hovering
				manualAdvance:false, //Force manual transitions
				captionOpacity:1, //Universal caption opacity
				beforeChange: function(){$('.nivo-caption').animate({bottom:'-110'},400,'easeInBack')},
				afterChange: function(){Cufon.refresh();$('.nivo-caption').animate({bottom:'-20'},400,'easeOutBack')},
				slideshowEnd: function(){} //Triggers after all slides have been shown
			});
		   Cufon.refresh();
		});
		</script>
	</body>
</html>