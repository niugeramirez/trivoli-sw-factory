<% Option Explicit %>
<!--#include virtual="/rhprox2/shared/inc/sec.inc"-->
<!--#include virtual="/rhprox2/shared/inc/const.inc"-->
<!--#include virtual="/rhprox2/shared/db/conn_db.inc"-->
<!--
Archivo     : rep_auditoria_sup_00.asp
Autor       : JMH
Creacion    : 19/01/2005
Descripcion : Reporte de Auditoría
Modificacion:
-->
<%
Dim l_rs
Dim l_sql

Dim l_salida
Dim l_tipoModelo

l_salida = request("salida")
l_tipoModelo = request("modelo")

if l_salida = "" then
   l_salida = "rep_libro_ley_liq_04"
end if

if l_tipoModelo = "" then
   l_tipoModelo = "1"
end if

%>
<html>
<head>
<link href="/rhprox2/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title>Auditoría - RHPro &reg;</title>
<script src="/rhprox2/shared/js/fn_windows.js"></script>
<script src="/rhprox2/shared/js/fn_confirm.js"></script>
<script src="/rhprox2/shared/js/fn_ayuda.js"></script>
<script src="/rhprox2/shared/js/fn_ay_generica.js"></script>
<SCRIPT SRC="/rhprox2/shared/js/menu_def.js"></SCRIPT>
<script src="/rhprox2/shared/js/fn_fechas.js"></script>
<script src="/rhprox2/shared/js/fn_help_emp.js"></script>
<script src="/rhprox2/shared/js/fn_buscar_emp.js"></script>
<script>
var filtro="";
var opant = new String("");
var titulofiltro = new String("");
var tiposalida = 0;

function Imprimir(){
	parent.frames.ifrm.focus();
	window.print();
}

function agregarbprocnro(bprocnro){
	filtro= filtro + '&bprcnro='+ bprocnro;
}

function Filtro(tiposal){ 
	tiposalida = tiposal;
	// Siempre 9 parametros
	// opant-legal-concepto-acumulador-variosconceptos-varios acumuladores-nombreconcepto-nombreacumulador
	abrirVentana("reporte_filtro_auditoria_sup.asp?opant="+opant.valueOf(),'', 450, 430);
}

// 18 parametros siempre:
// tex, desde, hasta, aprobado, procesos,t1,e1,t2,e2,t3,e3, fecha, empresa, Conceptos, Acumuladores, desde, hasta, orden
function Filtrar(jsNuevo,taccion,accion,tusuario,usuario,fechadesde,fechahasta){

	var destino;
	filtro="filtro="+jsNuevo
					+ '&taccion='+taccion+'&accion='+accion
					+ '&tusuario='+tusuario+'&usuario='+usuario
					+ '&fechadesde='+fechadesde+'&fechahasta='+fechahasta
                    +'&tfiltro='+titulofiltro;	
	if ((jsNuevo != null) && (jsNuevo != "")) {
	   setTimeout("lanzarProceso()", 1000);
	} 

}

function lanzarProceso(){
    var destino;
	var arr;
  	//var pagina = showModalDialog('rep_libro_ley_liq_06.asp', '','dialogWidth:25;dialogHeight:13;help: 0; status: 0; resizable:0; center:1;scroll:0');

    //if (pagina != ''){
	  // arr = pagina.split('@');
  	   destino = "rep_auditoria_sup_01.asp?" ;
	   document.ifrm.location = destino+filtro;
	//}
}

function cambioLibroLey(bpronro){
   document.datos.bpronro.value = bpronro;

   if (bpronro != ''){
       document.ifrm.location = "<%= l_salida%>.asp?bpronro=" + bpronro;
   }else{
       document.ifrm.location = 'blanc.html';
   }	
}

function actualizar(bpronro){
  document.ifrmlibroley.location = 'combo_libro_ley_liq_00.asp?ancho=720&bpronro=' + bpronro;  
}

function baja(){
  if (document.datos.bpronro.value == ""){
    alert('Debe seleccionar un registro histórico.');
  }else{
    if (confirm('¿Desea borrar el histórico seleccionado?.')){
        abrirVentanaH('rep_libro_ley_liq_05.asp?bpronro=' + document.datos.bpronro.value ,'',100,100);
	}
  }
}

</script>

</head>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name=datos>
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
          &nbsp;		  
          <% call MostrarBoton ("sidebtnSHW", "Javascript:baja();","Baja") %>		
		  &nbsp;  
          <a class=sidebtnSHW href="Javascript:Imprimir()">Imprimir</a>			  
          &nbsp;
  	      <a class=sidebtnSHW href="Javascript:abrirVentana('../gti/procesamiento_de_gti_04.asp','',650,250);">Monitor</a>
		  &nbsp;&nbsp;&nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
		<tr>
		   <td width="25%">
		     <br>
		   </td>
		   <td width="50%" align="left" valign="bottom" colspan="1">
		     <b>Hist&oacute;ricos:</b>&nbsp;<br>	 
		     <iframe frameborder="0" name="ifrmlibroley" scrolling="No" src="combo_hist_auditoria_sup_00.asp?ancho=720" width="720" height="22"></iframe>
		   </td>
		   <td width="25%">
		     <br>
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
