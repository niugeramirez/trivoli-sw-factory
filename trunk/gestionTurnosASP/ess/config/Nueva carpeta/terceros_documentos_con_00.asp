<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: terceros_documentos_con_00.asp
'Descripción: Consulta de documentos asociados a tipos de terceros
'Autor : Lisandro Moro
'Fecha: 17/02/2005

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

on error goto 0

' Filtro
Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
Dim l_Orden      ' Son las etiquetas que aparecen en el orden
Dim l_CamposOr   ' Son los campos para el orden

' Filtro
l_etiquetas = "Descripción:"
l_Campos    = "tipdocdes"
l_Tipos     = "T"

' Orden
l_Orden     = "Descripción:;Estado:"
l_CamposOr  = "tipdocdes;Oblig"

Dim l_rs
Dim l_sql
Dim l_tercero
Dim l_descripcion
Dim l_tipternro
Dim l_ternro

l_tipternro = request.querystring("tipternro")
l_ternro = request.querystring("cabnro")
l_descripcion = request.querystring("descripcion")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT tipternro,tipterdes "  
l_sql = l_sql & " FROM tkt_tipotercero "
l_sql = l_sql & " WHERE tipternro = " & l_tipternro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_tercero = l_rs("tipterdes")
end if 

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Documentos - Ticket</title>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag){
	abrirVentana('../shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag){
	abrirVentana('../shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function param(){
	var setear = "cabnro=<%= l_ternro %>&tipternro=<%= l_tipternro %>";
	return setear;
}


function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("terceros_documentos_con_excel.asp?cabnro=<%= l_ternro %>&tipternro=<%= l_tipternro %>&desc=<%= "Documentos asociados a " & l_tercero %>&orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value) + '&descripcion=<%= l_descripcion %>','execl',250,150);
}

</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">Documentos</td>
          <td nowrap align="right" class="barra">
			<% 
			'tipternro	tipterdes
			'1			Vendedor
			'2			Camionero
			'3			Entregador/Recibidor
			'4			Empresa
			'5			Destinatario
			'6			Transportista
			'7			Corredor
			'8			Vendedor/Corredor
			'9			Entregador
			'10			Recibidor
			'11			Cuenta y Orden
		  	if ((l_tipternro = 3) OR (l_tipternro = 9) OR (l_tipternro = 10)) Then
				call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('terceros_documentos_con_02.asp?tipternro=" & l_tipternro & "&ternro=' + document.ifrm.datos.cabnro.value + '&tipdocnro=' + document.ifrm.datos.tipdocnro.value,'',400,135);","Modificar")
			end if
			%>
		  &nbsp;&nbsp;
		  <% call MostrarBoton ("sidebtnABM", "Javascript:llamadaexcel();","Excel")%>
		  &nbsp;&nbsp;
		  <a class=sidebtnSHW href="Javascript:orden('../../config/terceros_documentos_con_01.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('../../config/terceros_documentos_con_01.asp')">Filtro</a>
		  &nbsp;&nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
		<tr>
			<td colspan="2" align="center"><b><%= l_tercero %>&nbsp;:&nbsp;<input  size="60"   class="deshabinp" type="text" readonly value="<%= l_descripcion %>"></b></td>
		</tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe name="ifrm" src="terceros_documentos_con_01.asp?cabnro=<%= l_ternro %>&tipternro=<%= l_tipternro %>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
      </table>
</body>
</html>
