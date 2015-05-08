<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

	<title>Men� Vertical en acorde�n con jQuery</title>
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.3/jquery.min.js" type="text/javascript" charset="utf-8"></script>
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
		#menu ul a{background-color:#6594D1;}
		#menu ul a:hover{
			background-color:#FFF;
			color:#2961A9;
			text-shadow:none;
			-webkit-transition: color, background-color 0.2s linear;
		}
		ul{
			display:block;
			background-color:#2961A9;
			margin:0;
			padding:0;
			width:130px;
			list-style:none;
		}
		#menu ul{background-color:#6594D1;}
		#menu li ul {display:none;}
	</style>
</head>
 
<body>
<ul id="menu">
<li><a href="#">Menu 1</a>
	<ul>
		<li><a href="#">Submenu 1</a></li>
		<li><a href="#">Submenu 2</a></li>
		<li><a href="#">Submenu 3</a></li>
		<li><a href="#">Submenu 4</a></li>
	</ul>
</li>
<li><a href="#">Menu 2</a>
	<ul>
		<li><a href="#">Submenu 1</a></li>
		<li><a href="#">Submenu 2</a></li>
		<li><a href="#">Submenu 3</a></li>
		<li><a href="#">Submenu 4</a></li>
	</ul>
</li>
<li><a href="#">Menu 3</a>
	<ul>
		<li><a href="#">Submenu 1</a></li>
		<li><a href="#">Submenu 2</a></li>
		<li><a href="#">Submenu 3</a></li>
		<li><a href="#">Submenu 4</a></li>
	</ul>
</li>
<li><a href="#">Menu sin submenu</a></li>
</ul>
</body>
</html>
