<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
	"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
 
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
 
	<title>Menú Vertical en acordeón con jQuery</title>
	<script src="../../js/jquery.min.js" type="text/javascript" charset="utf-8"></script>
	<script type="text/javascript" charset="utf-8">
	$(function(){
		$('#menu li a').click(function(event){
			var elem = $(this).next();
			if(elem.is('ul')){
				event.preventDefault();
				$('#menu ul:visible').not(elem).slideUp();
				elem.slideToggle();
			}
		});
	});
	</script>
	<style type="text/css" media="screen">
		#menu{
			-moz-border-radius:5px;
			-webkit-border-radius:5px;
			border-radius:5px;
			-webkit-box-shadow:1px 1px 3px #888;
			-moz-box-shadow:1px 1px 3px #888;
		}
		#menu li{border-bottom:1px solid #FFF;}
		#menu ul li, #menu li:last-child{border:none}	
		a{
			display:block;
			color:#FFF;
			text-decoration:none;
			font-family:'Helvetica', Arial, sans-serif;
			font-size:13px;
			padding:3px 5px;
			text-shadow:1px 1px 1px #325179;
		}
		#menu a:hover{
			color:#F9B855;
			-webkit-transition: color 0.2s linear;
		}
		#menu ul a{background-color:#FFFFFF;color:#000000}
		#menu ul a:hover{
			background-color:#A3E4E1;
			color:#2961A9;
			text-shadow:none;
			-webkit-transition: color, background-color 0.2s linear;
		}
		ul{
			display:block;
			background-color:#2FA69F; 
			margin:0;
			padding:0;
			width:150px;
			list-style:none;
		}
		#menu ul{background-color:#6594D1;}
		#menu li ul {display:none;}
	</style>
</head>
 
<body bgcolor="#F1F2F2">
<ul id="menu">

<!--
<li><a href="#">Agenda</a>
	<ul>		
		<li><a target="ifrm"  href="../../config/Agenda_con_00_v2.asp">Agenda</a></li>
	</ul>
</li>

-->
<li><a href="#">Ventas</a>
	<ul>		
		<li><a target="ifrm"  href="../../config/clientes_con_00.asp">Clientes</a></li>
		<li><a target="ifrm"  href="../../config/ventas_con_00.asp">Ventas</a></li>
	</ul>
</li>
<li><a href="#">Compras</a>
	<ul>		
		<li><a target="ifrm"  href="../../config/proveedores_con_00.asp">Proveedores</a></li>
		<li><a target="ifrm"  href="../../config/compras_con_00.asp">Compras</a></li>
	</ul>
</li>

<li><a href="#">Caja</a>
	<ul>		
		<li><a target="ifrm"  href="../../config/cheques_con_00.asp">Cheques</a></li>
		<li><a target="ifrm"  href="../../config/cajamovimientos_con_00.asp">Movimientos Caja</a></li>
	</ul>
</li>

<li><a href="#">Reportes</a>
	<ul>
		<li><a target="ifrm"  href="">A definir</a></li>
		<!--
		<li><a target="ifrm"  href="../../reportes/rep_visitas_entre_fechas_rep_00.asp">Visitas entre Fechas</a></li>
		<li><a target="ifrm"  href="../../reportes/rep_pagos_por_medio_rep_00.asp">Pago entre Fechas</a></li>
		-->
	</ul>
</li>
<li><a href="#">Configuraci&oacute;n</a>
	<ul>				
		<li><a target="ifrm"  href="../../config/bancos_con_00.asp">Bancos</a></li>
		
	</ul>
</li>
<!--<li><a href="#">Menu sin submenu</a></li>-->
</ul>
</body>
</html>
