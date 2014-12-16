<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->

<% 
on error goto 0

'-------------------------------------------------------------------------------------------
'Archivo		: filtro_campos_02.asp
'Descripción : Permite realizar un filtrado 
'Autor		: Lic. Fernando Favre
'Fecha		: 01/2004
'Modificado	: 
'23/11/2004 - Alvaro Bayon - Validaciones para código y descripción
'---------------------------------------------------------------------------------------------
 
 Dim l_tipo
 Dim l_campo
 
 l_campo = request("campo")
 l_tipo  = request("tipo")
 
%>
<html>	
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Filtro - RHPro &reg;</title>
</head>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script>
function filtrartxt(){
 var filtro;
 var expReg
 	expReg = new RegExp("'","ig");
	for (i=0;i<=2;i++){
	    if (document.datos.orden[i].checked)
			var sel = document.datos.orden[i].value
	}
  	if (document.datos.texto.value == ""){
    	alert("Debe ingresar un texto")
		document.datos.texto.focus()
		}
  	else
	{
		document.datos.texto.value = document.datos.texto.value.replace(expReg, "");
		if (sel == "1")	
	  	    var filtro = "<%= l_campo %> LIKE '" + document.datos.texto.value + "%'"
		if (sel == "2")	
	  	    var filtro = "<%= l_campo %> LIKE '%" + document.datos.texto.value + "%'"
		if (sel == "3")	
	  	    var filtro = "(<%= l_campo %> = '" + document.datos.texto.value + "')"
		
		window.parent.ifrm_datos.filtro.value = filtro;
		window.parent.Filtrar();
    }
}

function filtrarnum(){
	for (i=0;i<=2;i++){
		if (document.datos.orden[i].checked)
			var sel = document.datos.orden[i].value
  	}
  	if (document.datos.texto.value == ""){
    	alert("Debe ingresar un valor")
		document.datos.texto.focus()
		}
  	else
    if (isNaN(document.datos.texto.value)) {
	 	alert("El valor debe ser numérico.");
		document.datos.texto.focus()
		}
	else
  	if (!validanumero(document.datos.texto,9,0)){
    	alert("Debe ingresar un valor de hasta 9 cifras")
		document.datos.texto.focus()
		}
    else{
		if (sel == "1")	
	  	    var txt = '<%= l_campo %> > ' + document.datos.texto.value;
		if (sel == "2")	
	  	    var txt = '<%= l_campo %> < ' + document.datos.texto.value;
	  	if (sel == "3")	
	  	    var txt = '<%= l_campo %> = ' + document.datos.texto.value;
			
		window.parent.ifrm_datos.filtro.value = txt;
		window.parent.Filtrar();
    }
}

function filtrarfec(){
	for (i=0;i<=2;i++){
	    if (document.datos.orden[i].checked)
	  		var sel = document.datos.orden[i].value
	}
	if (document.datos.texto.value == ""){
    	alert("Debe ingresar un valor")
		document.datos.texto.focus()
		}
	else
    	if (validarfecha(document.datos.texto)){
	  		switch (sel){
				case "1":	
		  	    	var filtro = '<%= l_campo %> < ' + consultafecha(document.datos.texto.value);
					break;
	  			case "2":	
  	    			var filtro = '<%= l_campo %> > ' + consultafecha(document.datos.texto.value);
					break;
				case "3":
				if (document.datos.entre.value == "")	
			    	alert("Debe ingresar un valor");
				else
					if (validarfecha(document.datos.entre))
				  	    var filtro = '(' + consultafecha(document.datos.entre.value) + ' <= <%= l_campo %>) AND (<%= l_campo %> <= ' + consultafecha(document.datos.texto.value) + ')';
			}
			window.parent.ifrm_datos.filtro.value = filtro;
			window.parent.Filtrar();
    	}
}

function cambia(){
	for (i=0;i<=2;i++){
	    if (document.datos.orden[i].checked)
	  		var sel = document.datos.orden[i].value
	}
	if (sel == "3"){
		document.datos.entre.className = "habinp";
		document.datos.entre.readOnly = false;	
		document.datos.entre.focus();	
	}
	else{
		document.datos.entre.className = "deshabinp";
		document.datos.entre.readOnly = true;	
		document.datos.texto.focus();	
	}
}

function Tecla(num){
	if (num==13)
		return false;
	return num;
}

</script>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript: <% if l_tipo <> "R" then %> document.datos.texto.focus(); <% end if %>">
<form name="datos" method="post">
<table cellspacing="0" cellpadding="0" cellspacing="0" border="0" height="100%">
<%
 select case l_tipo
	case "N":
 %>
	<tr>
    	<td align="right"><b>Mayor a:</b></td>
		<td><input type="Radio" name="orden" value="1" checked></td>
    	<td rowspan="3" align="center">
		<input onclick="javascript:filtrarnum();" type=button value="Filtrar" name="btnFiltrar" class="sidebtnABM" style="background-color: #4682B4; color: White; border: 1px solid black; cursor: hand; width: 70;"> 
		</td>
	</tr>
	<tr>
	    <td align="right"><b>Menor a:</b></td>
		<td><input type="Radio" name="orden" value="2"></td>
	</tr>
	<tr>
	    <td align="right"><b>Igual a:</b></td>
		<td><input type="Radio" name="orden" value="3"></td>
	</tr>
	<tr>
	    <td align="right"><b>Valor:</b></td>
		<td colspan="2"><input type="Text" name="texto" value="" onkeypress="return Tecla(event.keyCode);"></td>
	</tr>
<% 
 	case "T":
 %>
	<tr>
	    <td align="right"><b>Comienza con:</b></td>
		<td><input type="Radio" name="orden" value="1" checked></td>
    	<td rowspan="3" align="center">
		<input onclick="javascript:filtrartxt();" type=button value="Filtrar" name="btnFiltrar" class="sidebtnABM" style="background-color: #4682B4; color: White; border: 1px solid black; cursor: hand; width: 70;"> 
		</td>
	</tr>
	<tr>
	    <td align="right"><b>Contiene:</b></td>
		<td><input type="Radio" name="orden" value="2"></td>
	</tr>
	<tr>
	    <td align="right"><b>Igual a:</b></td>
		<td><input type="Radio" name="orden" value="3"></td>
	</tr>
	<tr>
	    <td align="right"><b>Texto:</b></td>
		<td colspan="2"><input type="Text" name="texto" value="" onkeypress="return Tecla(event.keyCode);"></td>
	</tr>
<%
	case "F":
%>
	<tr>
		<td align="right"><b>Anterior a:</b></td>
		<td><input type="Radio" name="orden" value="1" checked onclick="Javascript:cambia();"></td>
    	<td rowspan="2" align="center">
		<input onclick="javascript:filtrarfec();" type=button value="Filtrar" name="btnFiltrar" class="sidebtnABM" style="background-color: #4682B4; color: White; border: 1px solid black; cursor: hand; width: 70;"> 
		</td>
	</tr>
	<tr>
	    <td align="right"><b>Posterior a:</b></td>
		<td><input type="Radio" name="orden" value="2" onclick="Javascript:cambia();"></td>
	</tr>
	<tr>
	    <td align="right"><b>Entre el:</b></td>
		<td><input type="Radio" name="orden" value="3" onclick="Javascript:cambia();"></td>
		<td><input type="Text" name="entre" size="10" maxlength="10" value="" class="deshabinp" readonly onkeypress="return Tecla(event.keyCode);"></td>
	</tr>
	<tr>
	    <td align="right"><b>y el:</b></td>
		<td>&nbsp;</td>
		<td colspan="2"><input type="Text" name="texto" size="10" maxlength="10" value="" onkeypress="return Tecla(event.keyCode);"></td>
	</tr>
<%
	case "R":
%>
	<tr>
	    <td>&nbsp;</td>
	</tr>
	<tr>
	    <td>&nbsp;</td>
	</tr>
	<tr>
	    <td>&nbsp;</td>
	</tr>
	<tr>
	    <td>&nbsp;</td>
	</tr>
<%
 end select
%> 	
</table>
</form>
</body>
</html>