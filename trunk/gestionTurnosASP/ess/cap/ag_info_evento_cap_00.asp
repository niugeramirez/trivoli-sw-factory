<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: ag_info_evento_cap_02.asp
Descripción: 
Autor : Lisandro moro
Fecha: 29/03/2004
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
  l_etiquetas = "C&oacute;digo:;Descripción:"
  l_Campos    = "cap_modulo.modnro;cap_modulo.moddesabr"
  l_Tipos     = "N;T"

' Orden
  l_Orden     = "C&oacute;digo:;Descripción:"
  l_CamposOr  = "cap_modulo.modnro;cap_modulo.moddesabr"

Dim l_evenro
Dim l_evento
Dim l_curso
Dim l_rs
Dim l_sql
Dim l_portot
  
l_evenro	 = Request.QueryString("evenro")

 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT cap_evento.evecodext, cap_evento.evedesabr, cap_curso.curcodext, cap_curso.curdesabr, cap_evento.eveporasi "  
 l_sql = l_sql & " FROM cap_evento "
 l_sql = l_sql & " INNER JOIN cap_curso ON cap_curso.curnro = cap_evento.curnro " 
 l_sql = l_sql & " WHERE evenro = " & l_evenro
 rsOpen l_rs, cn, l_sql, 0 
 if not l_rs.eof then
 	l_evento = l_rs("evecodext") & " - " & l_rs("evedesabr")
 	l_curso  = l_rs("curcodext") & " - " & l_rs("curdesabr")
	l_portot = l_rs("eveporasi")
 end if 

%>
<html>
<head>
<link href="../<%= session("estilo")%>" rel="StyleSheet" type="text/css">
<title>Resúmen del Evento - Capacitación - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag)
{
  abrirVentana('ord_brow_param.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&param=evenro=<%=l_evenro%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag)
{
  abrirVentana('filtro_brow_param.asp?pagina='+pag+'&campos=<%= l_campos%>&param=evenro=<%=l_evenro%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("evento_modulos_cap_excel.asp?evenro= <%= l_evenro %>&orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}

function Refrescar() { 
		 document.ifrm.location.reload();
		 document.ifrm1.location.reload();
}



function Llamar(nro) { 
	nro = parseInt(nro);
	switch (nro) {
	   case 1 :
			document.ifrm.location.href="ag_cerrar_paraprodes_cap_00.asp?evenro= <%= l_evenro %> ";
			break;
	   case 2 :
		    document.ifrm.location.href="ag_cerrar_for_cap_00.asp?evenro= <%= l_evenro %>";		
			break;		
	   case 3 :
			document.ifrm.location.href="ag_cerrar_infadi_cap_00.asp?evenro= <%= l_evenro %>";		
			break;		
	   case 4 :
	   		document.ifrm.location.href=src="ag_cerrar_pargap_cap_00.asp?evenro= <%= l_evenro %>";   
			break;		
	   case 5 :
	   		document.ifrm.location.href="ag_cerrar_eval_cap_00.asp?evenro= <%= l_evenro %>";		
			break;		   
	  case 6 :
	   		document.ifrm.location.href="ag_asistencias_cap_03.asp?evenro= <%= l_evenro %>";		
			break;
	} 
}

</script>
</head>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" >
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr >
		  <td colspan="2">
		  	<table border="0" cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td class="th2" width="33%" height="25">&nbsp;</td>
						<td class="th2" align="center" width="34%"></td>
						<td class="th2" align="right" width="33%">
						<!--<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda&nbsp;</a>&nbsp;-->
						</td>
					</tr>
			</table>		  
		  </td>
        </tr>
		<tr>
			<td align="center">
				<table>
					<tr>
						<td width="50%">&nbsp;</td>
						<td align="right"><b>Evento: </b></td>
						<td align="left"  ><input border="0" readonly type="Text" class="deshabinp" style="width:350" value="<%= l_evento %>" ></td>
						<td width="50%">&nbsp;</td>
					</tr> 
					<tr>
						<td width="50%">&nbsp;</td>
						<td align="right"><b>Curso: </b></td>
						<td align="left"  ><input border="0" readonly type="Text" class="deshabinp" style="width:350" value="<%= l_curso %>" ></td>
						<td width="50%">&nbsp;</td>
					</tr> 
					<tr>
						<td width="50%">&nbsp;</td>
					  <td align="right" style="width:2px; align:right;"><b>Origen:</b></td>
			          <td align="left" style="width:1px;">
						<select name="opc" onchange="Javascript:Llamar(document.all.opc.value);">
							<option value=1 selected>Participantes
							<option value=2>Formadores
							<option value=3>Información Complementaria
							<option value=5>Evaluaciones
							<option value=6>Asistencias
						</select>
				      </td>
					  <td width="50%">&nbsp;</td>
			        </tr>
				</table>
			</td>
	    <tr valign="top" height="100%">
          <td colspan="2" style="">
      	  <iframe  frameborder="0" scrolling="Yes" name="ifrm" src="ag_cerrar_paraprodes_cap_00.asp?evenro=<%= l_evenro %>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
		
		

      </table>
</body>
</html>
