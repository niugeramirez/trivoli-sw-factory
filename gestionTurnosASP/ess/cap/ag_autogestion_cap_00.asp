<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: autogestion_cap_00.asp
Descripción: Ventana de Autogestion
Autor : Raul Chinestra
Fecha: 30/12/2004

-->
<% 
'on error goto 0 

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
  l_etiquetas = "Apellido:;Nombre:;Fecha:"
  l_Campos    = "terape;ternom;calfecha"
  l_Tipos     = "T;T;F"

' Orden
  l_Orden     = "Apellido:;Nombre:;Fecha:"
  l_CamposOr  = "terape;ternom;calfecha"

Dim l_evenro
Dim l_evento
Dim l_curso
Dim l_rs
Dim l_sql
Dim l_codigo
Dim rs9
Dim l_empleg
Dim l_ternro

l_ternro = l_ess_ternro
l_empleg = l_ess_empleg

%>

<html>
<head>
<link href="../<%= c_Estilo %>" rel="StyleSheet" type="text/css">
<title>Capacitación - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

function Llamar(nro) { 
	nro = parseInt(nro);
	switch (nro) {
	   case 1 :
			document.ifrm.location.href="ag_eventos_abiertos_cap_00.asp?ternro=<%= l_ternro %> ";
			break;
	   case 2 :
		    document.ifrm.location.href="ag_eventos_abiertos_alta_cap_00.asp?ternro=<%= l_ternro %>";	//ag_eventos_cerrados_cap_00.asp
			break;		
	   case 3 :
			document.ifrm.location.href="ag_solicitud_eventos_cap_00.asp?ternro= <%= l_ternro %>";		
			break;		
	} 
}

</script>
</head>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" bgcolor="F6ECCE" >
      <table border="0" cellpadding="0" cellspacing="0" height="100%"  width="100%">
	     <tr valign="top">
	      <td align="center" >
		  	<table border="0" cellpadding="0" cellspacing="0"  bordercolor="" width="10">
				<tr>
					<td align="left"  nowrap width="33%">
						<select name="opc" onchange="Javascript:Llamar(document.all.opc.value);">
							<option value=1 selected>Eventos<!-- Abiertos-->
							<option value=2>Oferta e inscripción<!--Eventos Cerrados-->
							<option value=3>Solicitud de Eventos
						</select>
					</td>					
				</tr>				
			</table>			
	      </td>		  
        </tr>
		
        <tr valign="top" height="100%">
          <td  style="" colspan="2" height="100%">
      	  <iframe frameborder="0" name="ifrm" scrolling="Yes" src="ag_eventos_abiertos_cap_00.asp?ternro=<%= l_ternro%>" width="100%" height="100%" ></iframe> 
	      </td>
        </tr>
      </table>
</body>
</html>
