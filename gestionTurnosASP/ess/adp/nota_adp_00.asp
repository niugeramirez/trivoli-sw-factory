<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'---------------------------------------------------------------------------------
'Archivo	: nota_adp_00.asp
'Descripción: ABM de notas_ter  viene de tercero
'Autor		: Claudia Cecilia Rossi
'Fecha		: 30-08-2003
'Modificado	: 08-11-2005 - Leticia A. - Adecuarlo para Autogestion.
'Modificado : 10/09/2006 Raul Chinestra - se agregó el campo Revisada a las Notas
'			  30-07-2007 - Diego Rosso - Se cambio le formato de la tabla.
'----------------------------------------------------------------------------------

' Variables
' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

' Filtro
  l_etiquetas = "Tipo Nota:;Remitente:;Fecha Ingr.:;Fecha Vto.:;Nota Abr.:"
  l_Campos    = "tiponota.tnodesabr;notremitente;notfecalta;notfecvenc;notas_ter.notatxt"
  l_Tipos     = "T;T;F;F;T"

' Orden
  l_Orden     = "Tipo Nota:;Remitente:;Fecha Ingr.:;Fecha Vto.:;Nota Abr.:"
  l_CamposOr  = "tiponota.tnodesabr;notremitente;notfecalta;notfecvenc;notas_ter.notatxt"

Dim l_empleg

l_empleg = request("empleg")
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Notas - Administraci&oacute;n de Personal - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>

<script>

function filtro(pag){
  abrirVentana('/ess/ess/shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag){
abrirVentana('/ess/ess/shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function param(){
	//chequear= "empleg=<%'=l_empleg %>&tnoconfidencial=<%'= l_tnoconfidencial %>";
	chequear= "empleg=<%=l_empleg %>";
	return chequear;
}

function Exportar(){
	var destino
	destino = "nota_adp_excel.asp?";
	destino = destino + "empleg=<%=l_empleg %>"  //&tnoconfidencial=<%'= l_tnoconfidencial %>";
	destino = destino + "&filtro=" + escape(document.ifrm.datos.filtro.value);
	destino = destino + "&orden="  + document.ifrm.datos.orden.value;
	abrirVentana(destino,'excel',350,250);
}

function Eliminar(){
 if (ifrm.jsSelRow != null){
   if ( ifrm.jsSelRow.cells(3).innerText.toUpperCase() == 'SI')
            alert('La Nota está Revisada. No se puede Eliminar.');
	else		
		eliminarRegistro(document.ifrm,'nota_adp_04.asp?notanro='+document.ifrm.datos.cabnro.value);
 }			
 else
    alert('Debe Seleccionar un Registro.');

}

function Modificar(){
 if (ifrm.jsSelRow != null){
   if ( ifrm.jsSelRow.cells(3).innerText.toUpperCase() == 'SI')
            alert('La Nota está Revisada. No se puede Modificar.');
	else		
		abrirVentanaVerif('nota_adp_02.asp?Tipo=M&empleg=<%= l_empleg %>&notanro='+document.ifrm.datos.cabnro.value,'',600,360);
 }			
 else
    alert('Debe Seleccionar un Registro.');

}

 	   
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<table border="0" cellpadding="0" cellspacing="0" height="5%">
<tr style="border-color :CadetBlue;">
	<th width="10%" align="left">Notas </th>
	<th width="90%" style="text-align:right">
	<% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('nota_adp_02.asp?Tipo=A&empleg="&l_empleg&"' ,'',600,360)","Alta") %>
    <a class=sidebtnABM href="Javascript:Eliminar();">Borrar</a> &nbsp;	
    <a class=sidebtnABM href="Javascript:Modificar();">Modifica</a> &nbsp;	
	&nbsp;
	<a class=sidebtnSHW href="Javascript:orden('/ess/ess/adp/nota_adp_01.asp');">Orden</a>
 	<a class=sidebtnSHW href="Javascript:filtro('/ess/ess/adp/nota_adp_01.asp');">Filtro</a>
 	&nbsp;
    <a class=sidebtnSHW href="Javascript:Exportar();">Excel</a> &nbsp;
</th>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" height="95%">
<tr valign="top" height="100%">
   <td colspan="3" style="">
   <iframe name="ifrm" src="nota_adp_01.asp?empleg=<%=l_empleg%>" width="100%" height="100%"></iframe> 
   </td>
</tr>
<tr>
	<td colspan="3" height="10"></td>
</tr>

</table>
</body>
</html>
