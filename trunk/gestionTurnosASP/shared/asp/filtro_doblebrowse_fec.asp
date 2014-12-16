<% Option Explicit %>
<!--
Archivo: filtro_doblebrowse_fec.asp
Descripción: Filtra items en el doble browse para tipo fecha
Autor: F. Favre  
Fecha: 10-03
Modificado:
-->
<%
 
 Dim l_campo 
 Dim l_lado
 'esto es para pasar el parametro a la funcion que formatea la fecha
 Dim l_base
 
 l_campo = request.querystring("campo")
 l_lado  = request.querystring("lado")
 l_base = Session("base")
%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Filtrar - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
function cambia(){
	for (i=0;i<=2;i++){
	    if (document.datos.orden[i].checked)
	  		var sel = document.datos.orden[i].value
	}
	if (sel == "3"){
		document.datos.entre.disabled = false;	
		document.datos.entre.className = "habinp";
		document.all.entreFechas.href = "Javascript:Ayuda_Fecha(document.datos.entre);";
		document.datos.entre.focus();	
	}
	else{
		document.datos.entre.disabled = true;
		document.datos.entre.className = "deshabinp";
		document.datos.entre.value = "";
		document.all.entreFechas.href = "#";
		document.datos.texto.focus();	
	}
}

function filtrar(){
	for (i=0;i<=2;i++){
	    if (document.datos.orden[i].checked)
	  		var sel = document.datos.orden[i].value
	}
	
	if (sel == "1"){
		if (document.datos.texto.value == ""){
	    	alert("Debe ingresar una fecha.");
			document.datos.texto.focus();
			return;
		}
	    if (!validarfecha(document.datos.texto)){
			document.datos.texto.focus();
			document.datos.texto.select();
			return;
		}
	   	var filtro = "menor(<%= l_campo %>, '" + document.datos.texto.value + "') && <%= l_campo %> != ''"
	}
	if (sel == "2"){
		if (document.datos.texto.value == ""){
	    	alert("Debe ingresar una fecha.");
			document.datos.texto.focus();
			return;
		}
	    if (!validarfecha(document.datos.texto)){
			document.datos.texto.focus();
			document.datos.texto.select();
			return;
		}
  		var filtro = "! menorque(<%= l_campo %>, '" + document.datos.texto.value + "') "
	}
	if (sel == "3"){
		if (document.datos.entre.value == ""){	
	    	alert("Debe ingresar un fecha.");
			document.datos.entre.focus();
			return;
		}
		if (!validarfecha(document.datos.entre)){
			document.datos.entre.focus();
			document.datos.entre.select();
			return;
		}
		if (document.datos.texto.value == ""){
	    	alert("Debe ingresar una fecha.");
			document.datos.texto.focus();
			return;
		}
	    if (!validarfecha(document.datos.texto)){
			document.datos.texto.focus();
			document.datos.texto.select();
			return;
		}
  	    var filtro = "(menorque('" + document.datos.entre.value + "', <%= l_campo %>)) && (menor(<%= l_campo %>, '" + document.datos.texto.value + "'))"
	}
	
	window.opener.Filtrar(<%= l_lado %>, filtro);
	window.close();  
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto){
	return w_in.showModalDialog(pagina,'', 'help:0;status:0;resizable:0;center:1;scroll:0;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

	if (jsFecha == null) txt.value = ''
	else txt.value = jsFecha;
}

window.resizeTo(300,190)

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript: document.datos.texto.focus();">
<form name="datos" method="post">
<table cellspacing="1" cellpadding="0" border="0" width="100%" height="100%">
 <tr>
    <td class="th2" colspan="3">Filtrar</td>
 </tr>
 <tr>
    <td align="right"><b>Anterior a:</b></td>
	<td colspan="2"><input type="Radio" name="orden" value="1" checked onclick="Javascript:cambia();"></td>
 </tr>
 <tr>
    <td align="right"><b>Posterior a:</b></td>
	<td colspan="2"><input type="Radio" name="orden" value="2" onclick="Javascript:cambia();"></td>
 </tr>
 <tr>
    <td align="right"><b>Entre el:</b></td>
	<td><input type="Radio" name="orden" value="3" onclick="Javascript:cambia();"></td>
	<td>
		<input type="Text" name="entre" value="" disabled="true" class="deshabinp" size="10" maxlength="10">
		<a name="entreFechas" href="#"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a></td>
	</td>
 </tr>
 <tr>
    <td align="right"><b>y el:</b></td>
	<td></td>
	<td>
		<input type="Text" name="texto" value="" size="10" maxlength="10">
		<a href="Javascript:Ayuda_Fecha(document.datos.texto)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a></td>
	</td>
 </tr>
 <tr>
    <td colspan="3" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:filtrar()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>		
	</td>
 </tr>
</table>
</form>
</body>
</html>
