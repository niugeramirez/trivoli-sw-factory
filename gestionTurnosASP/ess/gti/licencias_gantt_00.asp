<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/numero.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : licencias_gantt_00.asp
Descripcion    : Modulo que se encarga del control de las licencias
Creacion       : 30/08/2004
Autor          : Scarpa D.
Modificacion   : 07-10-2005 - Leticia A. - Adecuacion para que funcione desde Autogestion 
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

const Color_Feriado   = "#ccffff"
const Color_FinSemana = "#e0e0e0"
const Color_Aprobado  = "#ccffcc"
const Color_Pendiente = "#edbf6b"

' Variables
Dim l_ternro
Dim l_empleg
dim l_rs
dim l_sql

Set l_rs  = Server.CreateObject("ADODB.RecordSet")

' _____________________________________________________________________
' ----------- NO SE si se necesita ahora!! ----------------------------
dim leg
leg = Session("empleg")
if leg = "" then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	Response.End
end if

l_sql = "SELECT ternro FROM empleado WHERE empleado.empleg = " & leg
l_rs.Open l_sql, cn
if l_rs.eof then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	response.end
else 
  l_ternro = l_rs("ternro")
end if
l_rs.close
' ___________________________________________________________________________
'l_ternro = request.QueryString("ternro")

l_empleg = request("empleg")

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Licencia:;Apellido:;Fecha desde:;Fecha hasta:"
  l_Campos    = "tipdia.tddesc;empleado.terape;elfechadesde;elfechahasta"
  l_Tipos     = "T;T;F;F"

' Orden
  l_Orden     = "Licencia:;Apellido:;Fecha desde:;Fecha hasta:"
  l_CamposOr  = "tipdia.tddesc;empleado.terape;elfechadesde;elfechahasta"
%>
<script>

function cambio(){
  var a;
  if (correcto()){
     a = parseInt(document.all.anio.value,10);
	 if (a < 2004){
        document.all.anio.value	= '2004';
	 }
     document.ifrm1.location = "licencias_gantt_01.asp?empleg=<%=l_empleg%>&anio=" + document.all.anio.value;
     document.ifrm2.location = "licencias_gantt_02.asp?empleg=<%=l_empleg%>&anio=" + document.all.anio.value;	 
  }
}

function correcto(){
  var v = document.all.anio.value;
  
  if (v == ''){
	 alert('Debe ingresar un año.');
	 return 0;
  }else{
     if (isNaN(v)){
	    alert('El año debe ser numerico.');
	    return 0;		
	 }else{
     	return 1;
	 }
  }
}

function Sig(){
  if (correcto()){
     document.all.anio.value= (parseInt(document.all.anio.value,10) + 1);
     document.ifrm1.location = "licencias_gantt_01.asp?empleg=<%=l_empleg%>&anio=" + document.all.anio.value;
     document.ifrm2.location = "licencias_gantt_02.asp?empleg=<%=l_empleg%>&anio=" + document.all.anio.value;	 
  }
}

function Ant(){
  var a;
  if (correcto()){
     a = parseInt(document.all.anio.value,10);

	 if (a > 2004){
        document.all.anio.value= (parseInt(document.all.anio.value,10) - 1);  
        document.ifrm1.location = "licencias_gantt_01.asp?empleg=<%=l_empleg%>&anio=" + document.all.anio.value;
        document.ifrm2.location = "licencias_gantt_02.asp?empleg=<%=l_empleg%>&anio=" + document.all.anio.value;	 
	 }
  }
}

</script>

<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<title>Control Licencias - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<style>
.TH3
{
	background-color: #F7C521;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 9pt;
	FONT-WEIGHT: bold;
	padding : 0 2 0 0;
	width : auto;
} 
</style>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="95%">
        <tr style="border-color :CadetBlue;">
          <td colspan="1" align="left" class="th2">Control&nbsp;Licencias&nbsp;&nbsp;&nbsp;</td>
          <td colspan="1" align="left" class="th2">
		    <table cellpadding="0" cellspacing="0" width="10%">
			   <tr style="border-color :CadetBlue;">
				  <td>
		 	         &nbsp;&nbsp;&nbsp;<b>A&ntilde;o:</b>&nbsp;
				  </td>			   
			     <td align="top">
			        <a href="JavaScript:Ant();"><img align="absmiddle" src="/serviciolocal/shared/images/prev.jpg" alt="Anterior" border="0"></a>		  	 
				 </td>
				 <td align="top">
			        <input type="text" size="4" maxlength="4" name="anio" onchange="javascript:cambio();" value="<%= year(date())%>">	 
				 </td>
				 <td align="top">
			        <a href="JavaScript:Sig();"><img align="absmiddle" src="/serviciolocal/shared/images/next.jpg" alt="Siguiente" border="0"></a>	 
				 </td>
			   </tr>
			</table>
		  </td>
		  <td colspan="2" class="th2" width="80%">
		    &nbsp;
		  </td>
        </tr>
        <tr valign="top" height="100%">		  
          <td colspan="4" style="">
		     <table cellpadding="0" cellspacing="0" border="0" width="100%" height="100%">
			    <tr>
				   <td width="38%" align="right" style="padding-right:0px;padding-bottom:18px;">
			          <iframe name="ifrm1" scrolling="No" src="licencias_gantt_01.asp?empleg=<%=l_empleg%>&anio=<%= year(date())%>" width="100%" height="100%"></iframe> 		  	   
				   </td>
				   <td align="left" style="padding-left:0px;">
                      <iframe name="ifrm2" src="licencias_gantt_02.asp?empleg=<%=l_empleg%>&anio=<%= year(date())%>" width="100%" height="100%"></iframe> 		  				   
				   </td>
				</tr>			 
			 </table>
	      </td>		
        </tr>
      </table>
</body>
</html>
