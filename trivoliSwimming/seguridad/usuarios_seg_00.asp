<%Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: usuarios_seg_00.asp
'Descripción: ABM de usuarios 
'Autor: Alvaro Bayon
'Fecha: 21/02/2005
'Modificado:

' Variables
' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Id. Usuario:;Nombre Usuario:;Perfil:;Email:;"
  l_Campos    = "iduser;usrnombre;perfnom;usremail;"
  l_Tipos     = "T;T;T;T;"

' Orden
  l_Orden     = "Id. Usuario:;Nombre Usuario:;Perfil:;Email:;"
  l_CamposOr  = "iduser;usrnombre;perfnom;usremail"


%>
<html>
<head>
<link href="/trivoliSwimming/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Usuarios - Seguridad - Ticket</title>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script>

function orden(pag){
  abrirVentana('orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag){
  abrirVentana('filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("usuarios_seg_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name=datos>
<input type="hidden" name="seleccion">
</form>
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr style="border-color :CadetBlue;">
          <td colspan="2" align="left" class="barra">Usuarios</td>
          <td colspan="2" align="right" class="barra">
          <% 'call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('usuarios_seg_02.asp?Tipo=A','',550,350);","Alta")%>
		  <a class=sidebtnSHW href="Javascript:abrirVentana('usuarios_seg_02.asp?Tipo=A','',550,350);">Alta</a>
          <% call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'usuarios_seg_04.asp?iduser=' + document.ifrm.datos.cabnro.value);","Baja")%>
          <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('usuarios_seg_02.asp?Tipo=M&iduser=' + document.ifrm.datos.cabnro.value,'',550,350);","Modifica")%>
		  &nbsp;&nbsp;  
          <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('permisos_seg_00.asp?Tipo=M&iduser=' + document.ifrm.datos.cabnro.value,'',640,520);","Permisos")%>
		  &nbsp;&nbsp;
		  <% call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
		  &nbsp;&nbsp;
		  <a class=sidebtnSHW href="Javascript:orden('usuarios_seg_01.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('usuarios_seg_01.asp')">Filtro</a>
		  &nbsp;&nbsp;&nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="4" style="">
      	  <iframe name="ifrm" src="usuarios_seg_01.asp" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
        <tr>
          <td colspan="4" height="10"></td>
        </tr>
      </table>
</body>
</html>
<SCRIPT SRC="/trivoliSwimming/shared/js/menu_op.js"></SCRIPT>