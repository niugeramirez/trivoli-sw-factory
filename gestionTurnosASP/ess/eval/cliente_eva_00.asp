<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'================================================================================
'Archivo		: cliente_eva_00.asp
'Descripción	: Abm de Clientes
'Autor			: CCRossi
'Fecha			: 13-12-2004
'Modificado		: 
'================================================================================

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
  l_etiquetas = "C&oacute;digo:;C&oacute;d.Ext.:;Raz&oacute;n Social:;"
  l_Campos    = "evaclinro;evaclicodext;evaclinom"
  l_Tipos     = "N;T;T"

' Orden
  l_Orden     = "C&oacute;digo:;C&oacute;d.Ext.:;Raz&oacute;n Social:;"
  l_CamposOr  = "evaclinro;evaclicodext;evaclinom"

%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Clientes - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag)
{
  abrirVentana('orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag)
{
  abrirVentana('filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("cliente_eva_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}

</script>
</head>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="th2">Clientes</td>
          <td nowrap align="right" class="th2">
          <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('cliente_eva_02.asp?Tipo=A','',500,150);","Alta")%>
          <% call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'cliente_eva_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
          <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('cliente_eva_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',500,150);","Modifica")%>
          <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('engagement_eva_00.asp?Tipo=M&evaclinro=' + document.ifrm.datos.cabnro.value,'',550,250);","Engagements")%>
		  &nbsp;
		  <% call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
		  <a class=sidebtnSHW href="Javascript:orden('cliente_eva_01.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('cliente_eva_01.asp')">Filtro</a>
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="">
      	  <iframe name="ifrm" src="cliente_eva_01.asp" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
		<tr>
			<td colspan="2" height="20">
			</td>
		</tr>
      </table>
</body>
</html>
