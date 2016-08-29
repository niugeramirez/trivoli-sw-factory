<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/ess/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
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
  l_etiquetas = "Descripción:;"
  l_Campos    = "agedes"
  l_Tipos     = "T;"

' Orden
  l_Orden     = "Descripción:;"
  l_CamposOr  = "agedes"

  Dim l_rs
  Dim l_sql
  Dim l_idpracticarealizada
  Dim l_apellido
  Dim l_nro
  Dim l_os
  Dim l_practica
  
  l_idpracticarealizada = request("cabnro")
  
  

	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT  clientespacientes.apellido, clientespacientes.nombre, clientespacientes.nrohistoriaclinica , obrassociales.descripcion , practicas.descripcion practica"
	l_sql = l_sql & " FROM practicasrealizadas "
	l_sql = l_sql & " INNER JOIN visitas ON visitas.id = practicasrealizadas.idvisita "
	l_sql = l_sql & " INNER JOIN clientespacientes ON clientespacientes.id = visitas.idpaciente "
	l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
	l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = practicasrealizadas.idpractica "
	l_sql = l_sql & " where practicasrealizadas.id = " & l_idpracticarealizada 
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then  
		l_apellido = l_rs("apellido") & " " & l_rs("nombre")
		l_nro = l_rs("nrohistoriaclinica")
		l_os = l_rs("descripcion")
		l_practica = l_rs("practica")
	end if
  

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script src="js_pantallas/pagos.js"></script>

<!--	VENTANAS MODALES        -->
<!-- <script src="../js/ventanas_modales_custom_V2.js"></script>-->

<script>




$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlertPagos"									//id_dialogAlert
														);
								inicializar_dialogConfirmDelete(	"dialogConfirmDeletePagos"				//id_dialogConfirmDelete
																	,"pagosV2_con_04.asp"				//url_baja
																	,"dialogAlertPagos"						//id_dialogAlert
																	,"detalle_01_Pagos"						//id_form_datos
																	,"ifrm_pagos"								//id_ifrm_form_datos
																	,null //window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialogPagos" 										//id_dialog
														,"pagosV2_con_06.asp"							//url_valid_06
														,"pagosV2_con_03.asp"							//url_AM
														,"dialogAlertPagos"									//id_dialogAlert	
														,"datos_02_pagos"										//id_form_datos		
														,null //window.parent.ifrm.location					//location_reload
														,Validaciones_locales_pagos							//funcion_Validaciones_locales	
														,"ifrm_pagos"											//id_ifrm_form_datos														
														); 															
							});
</script>
<!--	FIN VENTANAS MODALES    -->

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
	<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
		<tr style="border-color :CadetBlue;">
		  <td align="left" class="barra">&nbsp;</td>
		  <td nowrap align="right" class="barra">
			<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialogPagos','pagosV2_con_02.asp?Tipo=A&idpracticarealizada=<%=l_idpracticarealizada%>',520,350)"><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Agregar Pago"></a>    
		  </td>
		</tr>
		<tr>
			<td align="right"><b>Paciente: </b></td>
			<td><input  type="text" name="legape" size="41" maxlength="21" value="<%= l_apellido %>" readonly   class="deshabinp" >
			</td>
		</tr>		
		<tr>
			<td align="right"><b>Nro. Historia Cl&iacute;nica: </b></td>
			<td><input  type="text" name="legape" size="41" maxlength="21" value="<%= l_nro %>" readonly   class="deshabinp">
			</td>
		</tr>		
		<tr>
			<td align="right"><b>Obra Social: </b></td>
			<td><input  type="text" name="legape" size="41" maxlength="21" value="<%= l_os %>" readonly   class="deshabinp">
			</td>
		</tr>	
		<tr>
			<td align="right"><b>Practica: </b></td>
			<td><input  type="text" name="legape" size="41" maxlength="21" value="<%= l_practica %>" readonly   class="deshabinp">
			</td>
		</tr>						
		<tr valign="top" height="100%">
			<td colspan="2" style="" width="100%">
				<iframe scrolling="Yes" name="ifrm_pagos"  id="ifrm_pagos"src="pagosV2_con_01.asp?idpracticarealizada=<%= l_idpracticarealizada  %>" width="100%" height="100%"></iframe> 
			</td>
		</tr>		
	</table>
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialogPagos" title="Pago"> 			</div>	  
				
		<div id="dialogAlertPagos" title="Mensaje">				</div>	

		<div id="dialogConfirmDeletePagos" title="Consulta">		</div>			
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->			  
</body>
</html>
