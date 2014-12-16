<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<html>
<head>
<link href="<%= c_estilo %>" rel="StyleSheet" type="text/css">
<title>Principal</title>
</head>
<script>
function opciones(boton){
	switch (parseInt(boton)){
		case 1:
			document.all.principal.src='personal.asp';
			break;
		case 2:
			document.all.principal.src='cap/ag_autogestion_cap_00.asp';
			break;
		case 3:
			document.all.principal.src='personal.asp';
			break;
		case 4:
			document.all.principal.src='personal.asp';
			break;
	}
}

var jsSelRow = null;
function Deseleccionar(fila){
	fila.className = "boton";
}
function Seleccionar(fila){
	if (jsSelRow != null){
	Deseleccionar(jsSelRow);
	 }
	fila.className = "botonsel";
	jsSelRow = fila;
}
</script>
<body class="indexprincipal">
	<table class="Tprincipal" cellpadding="0" cellspacing="0">
 		<tr>
			<td class="barmenu">
				Juan Jose Hernandez
				<a href="javascript:opciones(1);" class="boton" onClick="Seleccionar(this);">Personal</a>
				<a href="javascript:opciones(2);" class="boton" onClick="Seleccionar(this);">Principal</a>
				<a href="javascript:opciones(3);" class="boton" onClick="Seleccionar(this);">Otros</a>
				<a href="#" class="botonDSB" onClick="">Opciones</a>
<!-- 				<div style="float:right; position:avsolute; overflow-x:20; ">
					<img src="shared/images/a_capacitacion.gif" height="30px;">
				</div>
 -->
			</td>
		</tr>
		<tr>
			<td class="tdprincipal">
				<iframe name="principal" src="bienvenida.asp" frameborder="0" scrolling="no"></iframe>
			</td>
		</tr>
	</table>
</body>
<script>
		parent.document.all.centro.height = document.body.scrollHeight;
</script>
</html>
