<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<%
'--------------------------------------------------------------------------------------
'Archivo        : rep_auditoria_sup_00.asp
'Descripción    : Reporte - Auditoria
'Autor          : CCRossi
'Fecha Creacion : 28-04-2004
'Modificado     : 
'					21/07/2005 - Fapitalle N. - Agregar filtro por empleados
'					22/07/2005 - Fapitalle N. - Cambiar modo de pasaje de param (qs -> post)
'--------------------------------------------------------------------------------------

Dim l_salida

l_salida = "rep_auditoria_sup_04"

%>
<html>
<head>
<link href="/rhprox2/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title>Reporte - Auditor&iacute;a - Supervisor - RHPro &reg;</title>
<script src="/rhprox2/shared/js/fn_windows.js"></script>
<script src="/rhprox2/shared/js/fn_confirm.js"></script>
<script src="/rhprox2/shared/js/fn_ayuda.js"></script>
<script>
var filtro="";
var opant = new String("");
var titulofiltro = new String("");

function Imprimir(){
	parent.frames.ifrm.focus();
	window.print();
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else	
		abrirVentana("rep_auditoria_sup_02.asp?"+filtro,'execl',250,150);
}


function Filtro(sexcel){ 
	excel = sexcel;
	abrirVentana("rep_filtro_auditoria_sup.asp?opant="+opant.valueOf(),'', 450, 490);
}
function Filtrar(jsNuevo,acciones,acnro,usuarios,iduser,caudnro, fechadesde,fechahasta,orden,emptipo){
	var destino;
	
	filtro="filtro="+jsNuevo+ '&acciones='+acciones
					+ '&acnro='+acnro
					+ '&usuarios='+usuarios
					+ '&iduser='+iduser
					+ '&caudnro='+caudnro
					+ '&fechadesde='+fechadesde
					+ '&fechahasta='+fechahasta
					+ '&orden='+ orden +'&tfiltro='+titulofiltro
					+ '&emptipo='+ emptipo;

	if (emptipo == 1){
		var arrlist;
		var arremp;
		var i;
		var cad = "";
		
		arrlist = document.datos.empleados.value.split(",");
		document.datos.lista.value = "";
		for (i = 0; i < arrlist.length; i++){
			arremp = arrlist[i].split("@");
			if (arremp.length == 2){
				cad = cad + arremp[0] + ",";
			}
		}
		cad = cad.substr(0,cad.length-1);
		document.datos.lista.value = cad;
	}
	else
		document.datos.lista.value = document.datos.empleados.value;

	if ((jsNuevo != null) && (jsNuevo != "")) {
		if (excel)
			destino = "rep_auditoria_sup_02.asp?";
		else
			destino = "rep_auditoria_sup_01.asp?";
			
		if (excel)
			abrirVentana(destino+filtro,'execl',250,150);
		else{
			document.datos.target = 'ifrm';
			document.datos.action = destino+filtro;
			document.datos.submit();
		}
	}  
}

function actualizar(bpronro){
  document.ifrmauditoria.location = 'combo_hist_auditoria_sup_00.asp?ancho=720&bpronro=' + bpronro;  
}

function cambioAuditoria(bpronro){

   document.datos.bpronro.value = bpronro;

   if (bpronro != ''){
       document.ifrm.location = "<%= l_salida%>.asp?bpronro=" + bpronro;
   }else{
       document.ifrm.location = 'blanc.html';
   }	
}

function baja(){
	if (document.datos.bpronro.value == "")
    	alert('Debe seleccionar un registro histórico.');
	else{
    	if (confirm('¿Desea borrar el histórico seleccionado?.'))
			abrirVentanaH('rep_auditoria_sup_05.asp?bpronro=' + document.datos.bpronro.value + '&anchoselecthist=720','',100,100);
	}
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos" method="post">
<input type="hidden" name="empleados">
<input type="hidden" name="campos">
<input type="hidden" name="lista">

<input type="hidden" name="conceptonombre">
<input type="hidden" name="empresanombre" >
<input type="hidden" name="acumuladornombre" >
<input type="hidden" name="bpronro" value="">
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">Auditoría</td>
          <td nowrap align="right" colspan="2" class="barra">
		  &nbsp;&nbsp;&nbsp;
		  <a class=sidebtnSHW href="Javascript:Filtro(0)">Generar</a>
          <% call MostrarBoton ("sidebtnSHW", "Javascript:baja();","Baja") %>		
		  <a class=sidebtnSHW href="Javascript:Imprimir()">Imprimir</a>			  
          <a class=sidebtnSHW href="Javascript:abrirVentana('../gti/procesamiento_de_gti_04.asp','',650,250);">Monitor</a>
		  &nbsp;&nbsp;&nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
		<tr>
		   <td width="50%" align="center" valign="bottom" colspan="3">
		     <b>Hist&oacute;ricos:</b>&nbsp;<br>	 
		     <iframe frameborder="0" name="ifrmauditoria" scrolling="No" src="combo_hist_auditoria_sup_00.asp?ancho=720" width="720" height="22"></iframe>
		   </td>
		</tr>
        <tr valign="top" height="100%">
          <td colspan="3" align="center">
      	  <iframe name="ifrm" src="blanc.html" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
        <tr>
          <td colspan="3" height="10">
	      </td>
        </tr>
	</table>
</form>
</body>
</html>
