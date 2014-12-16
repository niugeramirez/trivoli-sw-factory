<% Option Explicit %>
<html>
<head>
<!--#include virtual="/serviciolocal/shared/inc/encrypt.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<%	
'modificado: 17/08/2007 - Lisandro Moro - Se modifico el titulo por (RH Pro X2)
'31/08/2007 - Martin Ferraro - Modificaciones para multiples bases de datos

	'Seteo la base default se debe configurar de acuerdo al default de la empresa
	'Idem lanzador
	Session("base") = "2" 'Actualmente conecta con X2
	Session.LCID = 11274
	'Session("Username") = Decrypt(c_seed1, c_x1)
	'Session("Password") = Decrypt(c_seed1, c_x2)
	
	'response.write "base " & Session("base") & "<br>"
	
%>
<link href="<%= c_Estilo %>" rel="StyleSheet" type="text/css">
<title>Cámara Portuaria y Marítima de Bahía Blanca</title>
</head>
<script src="shared/js/fn_util.js"></script>
<script>
var estadoEsMSS = 0;
var empleado = 0;
var accionMenu='#';
var ordenMenu='#';

function esMSS(){
   return (estadoEsMSS != 0);
}

function showtime() {
  var now = new Date();
  var hours = now.getHours();
  var minutes = now.getMinutes();
  var seconds = now.getSeconds()
  var timeValue = "" + hours
  timeValue += ((minutes < 10) ? ":0" : ":") + minutes;
  timeValue += ((seconds < 10) ? ":0" : ":") + seconds;
  //timeValue = now.getDate();
  document.all.tiempo.value = timeValue;
  timerID = setTimeout("showtime()",1000);
  timerRunning = true;
}

function accion(orden,accio){
	accionMenu=accio;
	ordenMenu=orden;		
	
	//alert(accio);
	
    if (esMSS()){
	    if (accio == ''){
	       document.all.centro.src='menuTop.asp?menu=mss&src=bienvenida.asp&orden=' + escape(orden) + '&empleg=' + empleado;
		}else{
		   document.all.centro.src='menuTop.asp?menu=mss&src=' + escape(accio) + '&orden=' + escape(orden) + '&empleg=' + empleado;
		}		
	}else{
	    if (accio == ''){
	       document.all.centro.src='menuTop.asp?menu=ess&src=bienvenida_44.asp&orden=' + escape(orden);	
		}else{
		   document.all.centro.src='menuTop.asp?menu=ess&src=' + escape(accio) + '&orden=' + escape(orden);
		}			
	}
}

function cambioMSS(){
    if (estadoEsMSS == 0){
		document.ifrmTerceros.location="terceros.asp";
	    estadoEsMSS = -1;
	}else{
	    document.ifrmTerceros.location="personas.asp";	
		ordenMenu = '';
		document.all.centro.src='menuTop.asp?menu=ess&src=bienvenida.asp&orden=&empleg=';	
	    estadoEsMSS = 0;	
	}
}

function accionTercero(empleg){
    empleado = empleg;
	
    if (accionMenu == ''){
       document.all.centro.src='menuTop.asp?menu=mss&src=bienvenida.asp&orden=' + escape(ordenMenu) + '&empleg=' + empleado;	
	}else{
	   document.all.centro.src='menuTop.asp?menu=mss&src=' + escape(accionMenu) + '&orden=' + escape(ordenMenu) + '&empleg=' + empleado;
	}		
}


</script>
<body onLoad="showtime();">
	<table class="index" border="0" cellspacing="0" cellspacing="0" align="center" height="100%">
		<tr>
			<td colspan="2" class="barmenu" height="1%">
				<div class="barsup">
					<%= date %> - 
					<input type="text" id="tiempo"  class="reloj" value="<%'= date %>" size="5">
				</div>
			</td>
		</tr>
		<tr>
			<td colspan="2" class="encabezado" height="3%">
				<table width="100%" cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td valign="top" align="center" width="100" class="tdlogo"> 
						
							<!--
							CAMARA PORTUARIA y MARITIMA <BR>							
							<img src="shared/images/raul.gif"  class="logo"><BR>
							DE BAHIA BLANCA							
							-->
							
						</td> 
						
						<td class="personas" align="left">						
							<iframe style="height=60px;" name="ifrmTerceros" src="personas.asp" frameborder="0" scrolling="no"></iframe>							
							<!--<iframe style="height=100px;" name="ifrmTerceros" src="" frameborder="0" scrolling="no"></iframe>-->														
							 <!--<img src="shared/images/titulo.jpg"> -->
 						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td class="menu" width="18%" height="90%">
				<div class="menu">
					<iframe src="menu.asp" name="ifrmmenu" frameborder="0" scrolling="no"></iframe>
				</div>
			</td>
			<td width="82%" height="90%">
				<div class="centro">
					<iframe src="principal.asp" name="centro" frameborder="0" scrolling="No"></iframe>
				</div>
			</td>
		</tr>

		<tr>
			<td colspan="2" class="barmenu">
				<div class="barsup">
				&nbsp;
				</div>
			</td>
		</tr>

	</table>

</body>
</html>
