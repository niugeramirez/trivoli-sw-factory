<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<%
'================================================================================
'Archivo		: proyecto_eva_ag_02.asp
'Descripción	: Abm de Proyectos
'Autor			: CCRossi
'Fecha			: 30-08-2004
'Modificado		: 09-02-2005 - L. Amadio - filtro por cliente.
'				: 11-02-2005 - L. Amadio - filtro por  engagement (relacionado con cliente) - eliminar campo nombre.
'				: 21-02-2005 - L. Amadio - restricion fechas de proyecto en relación al periodo del evento.
'				: 	 		 - L.Amadio -  alta engagement - - 
'				: 11-03-2005 - L.A. -  Alta:proy revisor, lo inicializa con emp reporta --- Ahora pidieron que se sacara!
'================================================================================
on error goto 0

'Datos del formulario
Dim l_evaproynro 
Dim l_evaproynom 
Dim l_evaproydext
Dim l_evaproyfht 
Dim l_evaproyfdd 
Dim l_evaclinro	 
Dim l_evaengnro	 
Dim l_proygerente
Dim l_proysocio	 
Dim l_proyrevisor
Dim l_estrnro	 
Dim l_proyaux1	 
Dim l_proyaux2   
Dim l_evapernro  

Dim l_evaperdesde
dim l_evaperhasta

Dim l_evaclicodext
Dim l_evaclinom   
Dim l_evaengcodext
Dim l_evaengdesabr

Dim l_nombre
Dim l_socio
Dim l_gerente
Dim l_revisor
Dim l_auxiliar1
Dim l_auxiliar2

Dim l_tipo 
Dim l_ternro
Dim l_perfil
Dim l_sql 
Dim l_rs  
Dim l_rs1 
Dim i     
Dim l_revmodifica
Dim l_estado

l_tipo = request.querystring("tipo")
l_ternro = request.querystring("ternro") ' ternro del usuario logeado
l_perfil = request.querystring("perfil") ' perfil del usuario logeado



sub modificarRevisor(modifica)
	' si cambia el revisor --> ver si ya se genero la evaluacion (se crearon evadetevldor)
	'						--> si se crearon evadetevdr  no se puede actualizar el revisor.
if l_tipo <> "A" then
	l_sql = "SELECT evadetevldor.evacabnro "
	l_sql = l_sql & " FROM evaproyecto "
	l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaproynro = evaproyecto.evaproynro "
	l_sql = l_sql & " INNER JOIN evacab ON evacab.evaproynro = evaevento.evaproynro "
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
	l_sql = l_sql & " WHERE evaproyecto.evaproynro ="& l_evaproynro
	rsOpen l_rs, cn, l_sql, 0 
	if l_rs.eof then ' no hay evaluacin
		modifica="SI"
	else
		modifica="NO"
	end if
	l_rs.close 
else 
	modifica="NO"
end if

end sub

'================================= BODY ==============================================
%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Proyectos - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
String.prototype.trim = function() {
 // skip leading and trailing whitespace
 // and return everything in between
  var x=this;
  x=x.replace(/^\s*(.*)/, "$1");
  x=x.replace(/(.*?)\s*$/, "$1");
  return x;
}

function Blanquear(texto){
 var  aux;
 aux = replaceSubstring(texto,"'","");
 aux = replaceSubstring(aux,"´","");
 
 return aux;
}

function replaceSubstring(inputString, fromString, toString) {
   // Goes through the inputString and replaces every occurrence of fromString with toString
   var temp = inputString;
   if (fromString == "") {
      return inputString;
   }
   if (toString.indexOf(fromString) == -1) { // If the string being replaced is not a part of the replacement string (normal situation)
      while (temp.indexOf(fromString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(fromString));
         var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
         temp = toTheLeft + toString + toTheRight;
      }
   } else { // String being replaced is part of replacement string (like "+" being replaced with "++") - prevent an infinite loop
      var midStrings = new Array("~", "`", "_", "^", "#");
      var midStringLen = 1;
      var midString = "";
      // Find a string that doesn't exist in the inputString to be used
      // as an "inbetween" string
      while (midString == "") {
         for (var i=0; i < midStrings.length; i++) {
            var tempMidString = "";
            for (var j=0; j < midStringLen; j++) { tempMidString += midStrings[i]; }
            if (fromString.indexOf(tempMidString) == -1) {
               midString = tempMidString;
               i = midStrings.length + 1;
            }
         }
      } // Keep on going until we build an "inbetween" string that doesn't exist
      // Now go through and do two replaces - first, replace the "fromString" with the "inbetween" string
      while (temp.indexOf(fromString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(fromString));
         var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
         temp = toTheLeft + midString + toTheRight;
      }
      // Next, replace the "inbetween" string with the "toString"
      while (temp.indexOf(midString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(midString));
         var toTheRight = temp.substring(temp.indexOf(midString)+midString.length, temp.length);
         temp = toTheLeft + toString + toTheRight;
      }
   } // Ends the check to see if the string being replaced is part of the replacement string or not
   return temp; // Send the updated string back to the user
} // Ends the "replaceSubstring" function


function Ayuda_Fecha(txt){
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha; 
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto) {
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}


function Validar_Formulario(){
var fechas
var fechasper
  // alert(document.datos.evaperiodo.value);


	// fechas del periodo de evaluac.

if (document.datos.evaperiodo.value !== '') {
 	// alert(document.datos.evaperiodo[document.datos.evaperiodo.selectedIndex].value);
	fechas = document.datos.evaperiodo[document.datos.evaperiodo.selectedIndex].value;
	fechaper= fechas.split(":");
	document.datos.evapernro.value   = fechaper[0];
	document.datos.evaperdesde.value = fechaper[1];
	document.datos.evaperhasta.value = fechaper[2];
}
document.datos.evaproynom.value  = Blanquear(document.datos.evaproynom.value);
document.datos.evaproydext.value = Blanquear(document.datos.evaproydext.value);

//alert(document.datos.evaperdesde.value);
//alert(document.datos.evaproyfdd.value);
	//if (menorque(document.datos.evaproyfdd.value,document.datos.evaperdesde.value) || !(menorque(document.datos.evaproyfdd.value,document.datos.evaperhasta.value))){
		//alert("La Fecha Desde debe estar en el periodo ...");
	//	document.datos.evaproyfdd.focus();
//	}

//return;

	
	if (document.datos.evaclinro.value == "" || (document.datos.evaclinro.value==0)) {
		alert("Seleccione un Cliente.");
		document.datos.evaclicodext.focus();
	}	
	else
	if ((document.datos.evaengnro.value == "") || (document.datos.evaengnro.value==0)){
		alert("Seleccione un Engagement.");
		document.datos.evaengcodext.focus();
	}	
	else
		/* if (document.datos.evaproynom.value.trim() == "") {
				alert("Debe ingresar un Nombre de Proyecto.");
				document.datos.evaproynom.focus();
		} 	else
		if (document.datos.evaproynom.value.length>30) 	{
				alert("El Nombre del Proyecto no puede superar 30 caracteres.");
				document.datos.evaproynom.focus();
		}	else  
		*/
	if (document.datos.evaproydext.value.trim()== "") {
		alert("Debe ingresar una Descripción Extendida.");
		document.datos.evaproydext.focus();
	}	
	else
	if (document.datos.evaproydext.value.length>255) {
		alert("La Descripción Extendida del Proyecto no puede superar 255 caracteres.");
		document.datos.evaproydext.focus();
	}	
	else
	if ((document.datos.proysocio.value=='') || (document.datos.proysocio.value=='0')) 	{
		alert("Ingrese un Socio.");
		document.datos.proysocio.focus();
	}	
	else
	if ((document.datos.proygerente.value=='') || (document.datos.proygerente.value=='0')){
		alert("Ingrese un Gerente.");
		document.datos.proygerente.focus();
	}	
	else
	if ((document.datos.proyrevisor.value=='') || (document.datos.proyrevisor.value=='0')) {
		alert("Ingrese un Revisor.");
		document.datos.proyrevisor.focus();
	}	
	else
	if ((document.datos.evaperiodo.value =='') || (document.datos.evaperiodo.value=='0')){
		alert("Seleccione un Período.");
		document.datos.evaperiodo.focus();
	}	
	else
	if (document.datos.evaproyfdd.value==""){
		alert("Ingrese una fecha.");
		document.datos.evaproyfdd.focus();
	}
	else
	if (!validarfecha(document.datos.evaproyfdd)){ 
		document.datos.evaproyfdd.focus();
		return
	}
	else
	if (document.datos.evaproyfht.value==""){
		alert("Ingrese una fecha.");
		document.datos.evaproyfht.focus();
	}
	else
	if (!validarfecha(document.datos.evaproyfht)){
		document.datos.evaproyfht.focus();
		return
	}
	else
	if (menorque(document.datos.evaproyfdd.value,document.datos.evaperdesde.value) || !(menorque(document.datos.evaproyfdd.value,document.datos.evaperhasta.value))){
		alert("La Fecha Desde debe estar dentro del rango del período.");
		document.datos.evaproyfdd.focus();
	}
	else
	if (menorque(document.datos.evaproyfht.value,document.datos.evaperdesde.value) || !(menorque(document.datos.evaproyfht.value,document.datos.evaperhasta.value))) {
		alert("La Fecha Hasta debe estar dentro del rango del período.");
		document.datos.evaproyfht.focus();
	}
	else
	if (!(menorque(document.datos.evaproyfdd.value,document.datos.evaproyfht.value))){
		alert("La Fecha Desde debe ser menor o igual que la Fecha Hasta.");
		document.datos.evaproyfdd.focus();
	}
	else
	if ((document.datos.estrnro.value=='') || (document.datos.estrnro.value=='0')) {
		alert("Seleccione una Línea de Servicio.");
		document.datos.estrnro.focus();
	} else {
		var d=document.datos;
		document.valida.location="proyecto_eva_ag_06.asp?tipo=<%=l_tipo%>&evaproynro="+document.datos.evaproynro.value + "&evaproynom="+document.datos.evaproynom.value + "&evaclinro=" + d.evaclinro.value  + "&evaengnro="+ d.evaengnro.value  + "&proysocio="+d.proysocio.value + "&proygerente="+ d.proygerente.value + "&proyrevisor="+ d.proyrevisor.value; 
	}
}


function valido(){
  document.datos.submit();
}

function invalido(texto){
  alert(texto);
}


function Teclarev(num){
  if (num==13) {
  		buscarcliente();
		return false;
  }

  return num;
}


function buscarcliente(esto,campo1, campo2){

if (esto==""){
	//
} else {
	if (isNaN(esto)){
		esto = "";
		alert("El Código ingresado no es correcto.")
	} else	 {
		abrirVentanaH('nuevo_rev.asp?empleg='+esto+'&campo1='+campo1+'&campo2='+campo2,'',200,100);
	}
}
}


// XXXXXXXXXXXXXXXXXXXXXXXX
function cliente(hay){
	if (document.datos.evaclicodext.value == ""){
			alert("Primero debe seleccionar un Cliente.");
			document.datos.evaclicodext.focus();
			hay=false 
	} else {
		hay = true
	}
}


function buscardatos(esto,campo1,campo2,campo3,dato){
if (dato=='E'){
	if (document.datos.evaclicodext.value == ""){
		alert("Primero debe seleccionar un Cliente.");
		document.datos.evaengcodext.value = "";
		document.datos.evaengnro.value = "0";
		document.datos.evaclicodext.focus();
		return;
	}
} else {
	// alert(document.datos.evaclicodext.value);
	
	document.datos.evaengcodext.value = "";
	document.datos.evaengdesabr.value = "";
	document.datos.evaengnro.value = "0";
	
	if (document.datos.evaclicodext.value == ""){
		alert("El Código ingresado no es correcto.");
		
		document.datos.evaclinom.value= "";
		document.datos.evaclinro.value="0";
		document.datos.evaclicodext.focus();
		return;
	}

}

	
if ( esto==""){  // isNaN(esto)||
		esto == "";
		 /*document.datos.evaclicodext.value = "";
		document.datos.evaclinom.value = "";
		document.datos.evaclinro.value = "0"; */
		
		document.datos.evaengcodext.value = "";
		document.datos.evaengdesabr.value = "";
		document.datos.evaengnro.value = "0";
		
		// alert(document.datos.evaengnro.value);
		
		alert("El Código ingresado no es correcto.");
} else {

	abrirVentanaH('nuevo_clienteeng.asp?cabnro='+esto+'&evaclinro='+document.datos.evaclinro.value+'&campo1='+campo1+'&campo2='+campo2+'&campo3='+campo3+'&dato='+dato,'',200,100);
}  
}

function altaEngagement (){
if ((document.datos.evaclinro.value=="") || (document.datos.evaclinro.value=='0')) {
	alert ('Primero debe especificar un Cliente.'); 
} else {
	abrirVentana('engagement_eva_00.asp?Tipo=M&evaclinro='+ document.datos.evaclinro.value+ '&campo1=document.datos.evaengcodext&campo2=document.datos.evaengdesabr&campo3=document.datos.evaengnro&dato=E&llama=PROY','',550,250);
	
	/*
		campo1=document.datos.evaengcodext&campo2=document.datos.evaengdesabr&campo3=document.datos.evaengnro&evaclinro='+document.datos.evaclinro.value +'&dato=E
	*/
}

}


//function setearValor(nro, desde, hasta) {
	//document.datos.evaperdesde.value = desde;
	//alert(document.datos.evaperdesde.value);
	//document.datos.evaperhasta.value = hasta;
// }


</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
select Case l_tipo
	Case "A":
		l_estado=""
		l_evaproynom = "proyecto"
		l_evaproyfht = date()
		l_evaproyfdd = date()
		
	 	l_sql = "SELECT evapernro, evaperdesabr, evaperdesde, evaperhasta "
		l_sql = l_sql & " FROM evaperiodo "
		l_sql = l_sql & " WHERE evaperdesde <= "& cambiafecha(Date(),"","")  & " AND evaperhasta >= "& cambiafecha(Date(),"","")
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then 
			l_evapernro=l_rs("evapernro")
			l_evaperdesde = l_rs("evaperdesde") 
			l_evaperhasta = l_rs("evaperhasta") 
		end if
		'response.write l_sql
		l_rs.Close
		'l_sql = "SELECT empleado.empreporta, reporta.empleg"
		'l_sql = l_sql & ", reporta.terape, reporta.terape2, reporta.ternom, reporta.ternom2 "
		'l_sql = l_sql & " FROM empleado "
		'l_sql = l_sql & " INNER JOIN empleado reporta ON reporta.ternro = empleado.empreporta"
		'l_sql = l_sql & " WHERE empleado.ternro=" & l_ternro
		'rsOpen l_rs, cn, l_sql, 0
		'if not l_rs.eof then
			'l_proyrevisor = l_rs("empleg")
			'l_revisor     = l_rs("terape") & " " & l_rs("terape2") & ", "  & l_rs("ternom") & " "& l_rs("ternom2")
		'end if
		'l_rs.Close
	Case "M": 
		l_estado="readonly disabled"
		l_evaproynro = request.querystring("cabnro")
		l_sql = "SELECT evaproynom, evaproyfht, evaproyfdd, evaproydext, evapernro "
		l_sql = l_sql & " , evaproyecto.evaengnro, evaengcodext, evaengdesabr "
		l_sql = l_sql & " , evacliente.evaclinro, evaclicodext, evaclinom "
		l_sql = l_sql & " , proygerente,proysocio,proyrevisor,estrnro,proyaux1,proyaux2 "
		l_sql = l_sql & " FROM evaproyecto "
		l_sql = l_sql & " INNER JOIN evaengage ON evaengage.evaengnro= evaproyecto.evaengnro "
		l_sql = l_sql & " INNER JOIN evacliente ON evacliente.evaclinro = evaengage.evaclinro "
		l_sql = l_sql & " WHERE evaproynro = " & l_evaproynro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then 
			l_evaproynom	= l_rs("evaproynom")
			l_evapernro		= l_rs("evapernro") 
			l_evaproyfht	= l_rs("evaproyfht")
			l_evaproyfdd	= l_rs("evaproyfdd")
			l_evaengnro		= l_rs("evaengnro") 
			l_evaengcodext 	= l_rs("evaengcodext")
			l_evaengdesabr 	= l_rs("evaengdesabr")
			l_evaclinro		= l_rs("evaclinro")
			l_evaclicodext 	= l_rs("evaclicodext")
			l_evaclinom 	= l_rs("evaclinom")
			l_proygerente	= l_rs("proygerente")
			l_proysocio		= l_rs("proysocio")
			l_proyrevisor	= l_rs("proyrevisor")
			l_estrnro		= l_rs("estrnro") 
			l_proyaux1		= l_rs("proyaux1")
			l_proyaux2		= l_rs("proyaux2")
			l_evaproydext	= l_rs("evaproydext")
		end if 
		l_rs.Close
		
		 ' ___________________________________
		 ' BUSCAR LOS EMPLEG .....            
		 ' ___________________________________
		if trim(l_proysocio)<>"" and not isnull(l_proysocio) then
				l_sql = "SELECT empleg FROM empleado WHERE ternro = "& l_proysocio
				rsOpen l_rs, cn, l_sql, 0 
				if not l_rs.eof then
					l_proysocio= l_rs("empleg")
				end if
				l_rs.close
		else
				l_proysocio=""
		end if
		if trim(l_proygerente)<>"" and not isnull(l_proygerente) then
				l_sql = "SELECT empleg FROM empleado WHERE ternro = "& l_proygerente
				rsOpen l_rs, cn, l_sql, 0 
				if not l_rs.eof then
					l_proygerente= l_rs("empleg")
				end if
				l_rs.close
		else
				l_proygerente=""	
		end if
		
		if trim(l_proyrevisor)<>"" and not isnull(l_proyrevisor) then
				l_sql = "SELECT empleg FROM empleado WHERE ternro = "& l_proyrevisor
				rsOpen l_rs, cn, l_sql, 0 
				if not l_rs.eof then
					l_proyrevisor= l_rs("empleg")
				end if
				l_rs.close
		else
				l_proyrevisor=""
		end if

		if trim(l_proyaux1)<>"" and not isnull(l_proyaux1) then
				l_sql = "SELECT empleg FROM empleado WHERE ternro = "& l_proyaux1
				rsOpen l_rs, cn, l_sql, 0 
				if not l_rs.eof then
					l_proyaux1= l_rs("empleg")
				end if
				l_rs.close
		else
				l_proyaux1=""	
		end if
		
		if trim(l_proyaux2)<>"" and not isnull(l_proyaux2) then
				l_sql = "SELECT empleg FROM empleado WHERE ternro = "& l_proyaux2
				rsOpen l_rs, cn, l_sql, 0 
				if not l_rs.eof then
					l_proyaux2= l_rs("empleg")
				end if
				l_rs.close
		else
				l_proyaux2=""	
		end if
		
		' buscar fecha del periodo
		if trim(l_evapernro)<>"" and not isnull(l_evapernro) then
			l_sql = "SELECT evapernro, evaperdesde, evaperhasta "
			l_sql = l_sql & " FROM evaperiodo "
			l_sql = l_sql  & " WHERE evapernro= " & l_evapernro 
			rsOpen l_rs, cn, l_sql, 0 
			if not l_rs.eof then 
				l_evaperdesde = l_rs("evaperdesde") 
				l_evaperhasta = l_rs("evaperhasta") 
			end if 
			l_rs.close 
		else
			l_evaperdesde = "" 
			l_evaperhasta = "" 
		end if

end select


%>
<!--
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.evaproydext.focus();Habilitar();CargarProyectos();">
-->
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.evaproydext.focus();">
<form name="datos" action="proyecto_eva_ag_03.asp?tipo=<%= l_tipo %>" method="post" >

<input type="Hidden" name="ternro"	   value="<%= l_ternro %>">
<input type="Hidden" name="perfil"	   value="<%= l_perfil %>">
<input type="Hidden" name="evaproynro" value="<%=l_evaproynro%>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="5%">
<tr>
    <td class="th2">Datos del Proyecto</td>
	<td class="th2" align="right">		  
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
</table>

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="96%">
<tr>
	<td align="right"><b>Cliente:</b></td>
	<td colspan=3>
		<input type="text" <%=l_estado%> name="evaclicodext" value="<%=l_evaclicodext%>" onKeyPress="return Teclarev(event.keyCode)" onChange="buscardatos(this.value,'document.datos.evaclicodext','document.datos.evaclinom','document.datos.evaclinro','C');" size="8" class="rev" > 
		<%if l_tipo="A" then%>
		<a onclick="JavaScript:window.open('help_clienteengag_00.asp?campo1=document.datos.evaclicodext&campo2=document.datos.evaclinom&campo3=document.datos.evaclinro&dato=C','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=yes,width=700,height=326');" onmouseover="window.status='Buscar Cliente'" onmouseout="window.status=' '" style="cursor:hand;">
		<img id="link1" name="link1" align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Cliente" border="0">
		</a>
		<%end if%>
		<input <%=l_estado%> class="rev" name="evaclinom" value="<%=l_evaclinom%>" style="background : #e0e0de;" readonly type="text" size="35" maxlength="35">
		<input type="Hidden" name="evaclinro"  value="<%=l_evaclinro%>">
	</td>
</tr>
<tr>
	<td nowrap align="right" colspan="4">  <!-- class="barra"  -->
	<%if l_tipo="A" then%>
	<a class=sidebtnABM href="#" onclick="Javascript:altaEngagement();"> Engagement </a>
	<%end if %>
	</td>
</tr>

<tr>
	<td align="right"><b>Engagement:</b></td>
	<td colspan=3>
		<input type="text" <%=l_estado%> name="evaengcodext" value="<%=l_evaengcodext%>" onKeyPress="return Teclarev(event.keyCode)" onChange="buscardatos(this.value,'document.datos.evaengcodext','document.datos.evaengdesabr','document.datos.evaengnro','E');" size="8" class="rev" > 
		<%if l_tipo="A" then%>
		<a onclick="JavaScript: window.open('help_clienteengag_00.asp?campo1=document.datos.evaengcodext&campo2=document.datos.evaengdesabr&campo3=document.datos.evaengnro&evaclinro='+document.datos.evaclinro.value +'&dato=E','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=yes,width=700,height=326');" onmouseover="window.status='Buscar Engagement'" onmouseout="window.status=' '" style="cursor:hand;">
		<img id="link1" name="link1" align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Engagement" border="0">
		</a>
		<%end if%>
		<input <%=l_estado%> class="rev" name="evaengdesabr" value="<%=l_evaengdesabr%>" style="background : #e0e0de;" readonly type="text" size="35" maxlength="80" >
		<input type="Hidden" name="evaengnro"  value="<%=l_evaengnro%>">

	</td>
</tr>
<tr>
    <td align="right"><b>Descripci&oacute;n:</b></td>
	<td colspan=3>
	<textarea name="evaproydext" rows="5" cols="40" maxlength="200"><%= l_evaproydext %></textarea>
	<input type="hidden" name="evaproynom" size="31" maxlength="30" value="<%=l_evaproynom %>">
	</td>
</tr>
<tr>
	<td align="right"><b>Socio:</b></td>
	<td colspan=3>
		<input type="text" <%=l_estado%> name="proysocio" value="<%=l_proysocio%>" onKeyPress="return Teclarev(event.keyCode)" onChange="buscarcliente(this.value,'document.datos.proysocio','document.datos.socio');" size="8" class="rev" > 
		<%if l_tipo="A" then%>
		<a onclick="JavaScript:window.open('help_emp_00.asp?campo1=document.datos.proysocio&campo2=document.datos.socio','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=yes,width=700,height=400');" onmouseover="window.status='Buscar Socio por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
		<img id="link1" name="link1" align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Socio" border="0">
		</a>
		<%end if%>
		<input <%=l_estado%> class="rev" name="socio" value="<%=l_socio%>" style="background : #e0e0de;" readonly type="text" size="35" maxlength="35" >

		<% if not (l_proysocio="" or l_proysocio="0") then %>
			<script>
				buscarcliente(document.datos.proysocio.value,'document.datos.proysocio','document.datos.socio')
			</script>
		<% end if %>			
	</td>
</tr>
<tr>
	<td align="right"><b>Gerente:</b></td>
	<td colspan=3>
		<input <%=l_estado%> type="text" name="proygerente" value="<%=l_proygerente%>" onKeyPress="return Teclarev(event.keyCode)" onChange="buscarcliente(this.value,'document.datos.proygerente','document.datos.gerente');" size="8" class="rev" > 
		<%if l_tipo="A" then%>
		<a onclick="JavaScript:window.open('help_emp_00.asp?campo1=document.datos.proygerente&campo2=document.datos.gerente','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=yes,width=700,height=400');" onmouseover="window.status='Buscar Gerentes por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
		<img id="link2" name="link2" align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Gerentes" border="0">
		</a>
		<%end if%>
		<input <%=l_estado%> class="rev" name="gerente" value="<%=l_gerente%>" style="background : #e0e0de;" readonly type="text" size="35" maxlength="35" >
		
		<% if not (l_proygerente="" or l_proygerente="0") then %>
			<script>
				buscarcliente(document.datos.proygerente.value,'document.datos.proygerente','document.datos.gerente')
			</script>
		<% end if %>			
	</td>
</tr>
<tr>
	<td align="right"><b>Revisor:</b></td>
	<td colspan=3>
		<% modificarRevisor l_revmodifica %>
		<%if (l_perfil<>"empleado" and l_revmodifica="SI") then %>
			<input type="text" name="proyrevisor" value="<%=l_proyrevisor%>" onKeyPress="return Teclarev(event.keyCode)" onChange="buscarcliente(this.value,'document.datos.proyrevisor','document.datos.revisor');" size="8" class="rev"> 
		<%else %>
			<input <%=l_estado%> type="text" name="proyrevisor" value="<%=l_proyrevisor%>" onKeyPress="return Teclarev(event.keyCode)" onChange="buscarcliente(this.value,'document.datos.proyrevisor','document.datos.revisor');" size="8" class="rev"> 
		<%end if %>
			
		<%if (l_tipo="A" or l_revmodifica="SI") then %>
		 <a onclick="JavaScript:window.open('help_emp_00.asp?campo1=document.datos.proyrevisor&campo2=document.datos.revisor','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=yes,width=700,height=400');" onmouseover="window.status='Buscar Revisor por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
		 <img id="link3" name="link3" align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Clientes" border="0">
		 </a>
		<%  %>
		<%end if%>
		
		<input <%=l_estado%> class="rev" name="revisor" value="<%=l_revisor%>" style="background : #e0e0de;" readonly type="text" size="35" maxlength="35" >
		
		<%if not (l_proyrevisor="" or l_proyrevisor="0") then %>
			<script>
				buscarcliente(document.datos.proyrevisor.value,'document.datos.proyrevisor','document.datos.revisor')
			</script>
		<%end if %>	
	</td>
</tr>
<tr>
	<td align="right"><b>Revisor Aux1.:</b></td>
	<td colspan=3>
		<input <%=l_estado%> type="text" name="proyaux1" value="<%=l_proyaux1%>" onKeyPress="return Teclarev(event.keyCode)" onChange="buscarcliente(this.value,'document.datos.proyaux1','document.datos.auxiliar1');" size="8" class="rev" > 
		<%if l_tipo="A" then%>
		<!--
		<a onclick="JavaScript:window.open('help_emp_00.asp?campo1=document.datos.proyrevisor&campo2=document.datos.revisor','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=yes,width=700,height=400');" onmouseover="window.status='Buscar Revisor por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
		-->
		<a onclick="JavaScript:window.open('help_emp_00.asp?campo1=document.datos.proyaux1&campo2=document.datos.auxiliar1','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=yes,width=700,height=400');" onmouseover="window.status='Buscar Revisor por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
		<img id="link3" name="link3" align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Clientes" border="0">
		</a>
		<%end if%>
		<input <%=l_estado%> class="rev" name="auxiliar1" value="<%=l_auxiliar1%>" style="background : #e0e0de;" readonly type="text" size="35" maxlength="35" >

		<% if not (l_proyaux1="" or l_proyaux1="0") then %>
			<script>
				buscarcliente(document.datos.proyaux1.value,'document.datos.proyaux1','document.datos.auxiliar1')
			</script>
		<% end if %>		
	</td>
</tr>
<tr>
	<td align="right"><b>Revisor Aux2.:</b></td>
	<td colspan=3>
		<input <%=l_estado%> type="text" name="proyaux2" value="<%=l_proyaux2%>" onKeyPress="return Teclarev(event.keyCode)" onChange="buscarcliente(this.value,'document.datos.proyaux2','document.datos.auxiliar2');" size="8" class="rev" > 
		<%if l_tipo="A" then%>
		<a onclick="JavaScript:window.open('help_emp_00.asp?campo1=document.datos.proyaux2&campo2=document.datos.auxiliar2','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=yes,width=700,height=400');" onmouseover="window.status='Buscar Revisor por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
		<img id="link3" name="link3" align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Clientes" border="0">
		</a>
		<%end if%>
		<input <%=l_estado%> class="rev" name="auxiliar2" value="<%=l_auxiliar2%>" style="background : #e0e0de;" readonly type="text" size="35" maxlength="35" >

		<% if not (l_proyaux2="" or l_proyaux2="0") then %>
			<script>
				buscarcliente(document.datos.proyaux2.value,'document.datos.proyaux2','document.datos.auxiliar2')
			</script>
		<% end if %>			
	</td>
</tr>
<tr>
	<%
	  Set l_rs = Server.CreateObject("ADODB.RecordSet")
	  l_sql = "SELECT estrnro,estrdabr"
	  l_sql = l_sql & " FROM estructura "
	  l_sql = l_sql & " WHERE  tenro= " & cdepartamento
	  rsOpen l_rs, cn, l_sql, 0 %>
	<td align="right"><b>L&iacute;nea de Servicio:</b></td>
	<td colspan=3>
		<select <%=l_estado%> name="estrnro">
		<option value="">< < Seleccione un Línea de Servicio > ></option>
		<%
		 do while not l_rs.eof%>
			<option value=<%=l_rs("estrnro")%>><%=l_rs("estrdabr")%></option>
		<%l_rs.MoveNext
		loop
		l_rs.Close
		set l_rs = nothing%>
		</select>
		<script>document.datos.estrnro.value='<%=l_estrnro%>'</script>		
	</td>
</tr>

<tr>
<%
	  Set l_rs = Server.CreateObject("ADODB.RecordSet")
	  l_sql = "SELECT evapernro, evaperdesabr, evaperdesde, evaperhasta "
	  l_sql = l_sql & " FROM evaperiodo  ORDER BY evaperdesde, evaperhasta "
	  rsOpen l_rs, cn, l_sql, 0 %>
	<td align="right"><b>Per&iacute;odo:</b></td>
	<td colspan=3>
		<select name="evaperiodo" <%=l_estado%>>
		<option value="">< < Seleccione un Período > ></option>
		<%
		 do while not l_rs.eof%>
			<option value="<%=l_rs("evapernro")%>:<%=l_rs("evaperdesde")%>:<%=l_rs("evaperhasta")%>"><%=l_rs("evaperdesabr")%> (<%=l_rs("evaperdesde")%> -- <%=l_rs("evaperhasta")%>)</option>
		<%l_rs.MoveNext
		loop
		l_rs.Close
		set l_rs = nothing%>
		</select>
		<input type="hidden" name="evapernro" value="<%=l_evapernro%>">
		<input type="hidden" name="evaperdesde" value="<%=l_evaperdesde%>">
		<input type="hidden" name="evaperhasta" value="<%=l_evaperhasta%>">

		<script>document.datos.evaperiodo.value='<%=l_evapernro%>:<%=l_evaperdesde%>:<%=l_evaperhasta%>'</script>
		<script>document.datos.evapernro.value='<%=l_evapernro%>'</script>
		
	</td>
</tr>
<tr>
    <td align="right"><b>Desde:</b></td>
	<td>
		<input type="text" name="evaproyfdd" size="10" maxlength="10" value="<%=l_evaproyfdd%>">
		<a href="Javascript:Ayuda_Fecha(document.datos.evaproyfdd)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
	</td>
    <td align="right"><b>Hasta:</b></td>
	<td>
		<input type="text" name="evaproyfht" size="10" maxlength="10" value="<%=l_evaproyfht%>">
		<a href="Javascript:Ayuda_Fecha(document.datos.evaproyfht)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
	</td>
</tr>
<tr height=42>
    <td  colspan="4" valign=top align="right" class="th2">
		<a class=sidebtnABM href="#" onclick="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>

	</td>
</tr> 	

</table>

<iframe name="valida" src="blanc.asp"  width="250px" height="150px"></iframe> <!--   style="visibility=hidden;"  -->
<iframe name="ifrm" style="visibility=hidden;" src="blanc.asp" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing


cn.Close
Set cn = Nothing
%>
</body>
</html>
