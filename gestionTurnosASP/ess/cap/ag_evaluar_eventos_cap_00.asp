<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: ag_evaluar_eventos_cap_00.asp
Descripción: Abm de Evaluacion de Eventos (Autogestión)
Autor : Raul Chinestra (listo)
Fecha: 21/06/2007
-->
<link href="../<%= c_estilo%>" rel="StyleSheet" type="text/css">
<% 

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "C&oacute;digo:;Descripción:"
  l_Campos    = "solnro;soldesabr"
  l_Tipos     = "N;T"

' Orden
  l_Orden     = "C&oacute;digo:;Descripción:"
  l_CamposOr  = "solnro;soldesabr"

Dim l_ternro
l_ternro  = request("ternro")

%>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<title>Evaluaciones - Capacitación - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag)
{
  abrirVentana('orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function cambiartipoevento()

{
 var indice;
 var valor;
	 indice = document.all.tipoevento.selectedIndex; 
	 valor  = document.all.tipoevento.options[indice].value; 
	 if (valor != 0)
 		{
	     document.ifrm.location = './ag_evaluar_satisfaccion_cap_01.asp?ternro=<%= l_ternro %>'
 		}
	 else
 		{
	 	 document.ifrm.location = './ag_evaluar_eventos_cap_01.asp?ternro=<%= l_ternro %>';	   
		} 
}   
function filtro(pag)
{
  abrirVentana('filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("contenidos_cap_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}

function Evaluar(){ 

 var destino;
 var indice;
 var valor;

	 indice = document.all.tipoevento.selectedIndex; 
	 valor  = document.all.tipoevento.options[indice].value; 
	 if (valor != 0)
 		{
    	 	 destino = 'ag_evaluar_satisfaccion_cap_02.asp?&ternro='+ document.datos.ternro.value+'&evenro='+ document.ifrm.datos.cabnro.value;	   
 		}
	 else
 		{
    	 	 destino = 'ag_evaluar_eventos_cap_02.asp?ttesnro='+document.ifrm.datos.ttesnro.value+'&ternro='+ document.datos.ternro.value+'&evenro='+ document.ifrm.datos.cabnro.value;	   
		} 

	if (document.ifrm.datos.cabnro.value=="0"){
		alert("Debe seleccionar un Evento");
	}
	else {
		abrirVentana(destino,'',780,580);
	}		
}

</script>
</head>

<form name="datos">
<input type=hidden name=ternro value="<%= l_ternro %>">
<input type=hidden name=ttestnro value="">
</form>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr style="border-color :CadetBlue;">
          <th align="left">
		  		Evaluación de Eventos:
			  	<select name="tipoevento" size="1" style="width:210;" onChange="Javascript:cambiartipoevento();">
			        <option value="0">Todas las evaluaciones</option> 				
			        <option value="1">Evaluaciones de Satisfacción</option> 				
				</select>
		  </th>
          <th nowrap align="right">
		  <a class=sidebtnABM href="Javascript:Evaluar();">Evaluar</a>          		  
		  &nbsp;&nbsp;
		  </th>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="">
      	  <iframe frameborder="0" name="ifrm"  scrolling="Yes" src="ag_evaluar_eventos_cap_01.asp?ternro=<%= l_ternro %>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
      </table>
</body>
<script>
window.document.body.scroll = "no";
</script>
</html>
