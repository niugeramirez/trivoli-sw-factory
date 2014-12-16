<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo     : rep_datos_emp_ess_00.asp
Autor       : GdeCos
Creacion    : 18/04/2005
Descripcion : Modulo que se encarga de generar el listado en el IFRAME del 00,
              datos del empleado.
Modificacion: 21/10/2005 - Leticia A. - Adecuarlo para Autogestion.
-->
<%
on error goto 0

Dim l_rs
Dim l_sql

Dim l_fecestr ' fecha para el his_estructura
Dim l_empleg

l_empleg = request("empleg")

 if l_fecestr = "" then 
 	l_fecestr = Date 
end if
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">

<title>Datos del Empleado - Autogestion - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
var indice = 1;
var ifrmListo = 0;

function Ayuda_Fecha(txt){
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function Imprimir(){
	ifrm.focus();
	window.print();
}


function ConfPagina(){
    var WebBrowser = '<OBJECT ID="WebBrowser' + indice + '" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
    document.body.insertAdjacentHTML('beforeEnd', WebBrowser);

    execScript("on error resume next: WebBrowser" + indice + ".ExecWB 8, -1", "VBScript");
	indice++;
}

function Mostrar(){
    document.datos.submit();
}


function actFechaEstr(opcion) {

	if (document.datos.orgfecha[0].checked) {
		document.datos.fechaestr.value = document.datos.fecestr.value;
	} else {
		document.datos.fechaestr.value = ""
	}

	if (opcion == 1){
		if (document.datos.orgfecha[0].checked)
			document.datos.submit();
	} else {
		document.datos.submit();
	}

}
</script>

</head>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos" action="rep_datos_emp_adp_01.asp?empleg=<%= l_empleg%>" method="post" target="ifrm">
<input type="hidden" name="posicion"  value="1">
<!--  EMPIEZA: Datos Empleado                    -->
<input type="hidden" name="fases"    	 value="-1">
<input type="hidden" name="organizacion" value="-1">
<input type="hidden" name="documentos"   value="-1">
<input type="hidden" name="domicilios"	 value="-1">
<input type="hidden" name="tipodom"    	 value="-1">
<!--  TERMINA: Datos Empleado					 -->
<!-- EMPIEZA: Datos Familiares                   -->
<input type="hidden" name="familiares" 	 value="-1">
<input type="hidden" name="parentesco"   value="0">
<!-- TERMINA: Datos Familiares 					 -->
<input type="hidden" name="fechaestr"  value="<%=l_fecestr %>">

    <table border="0" cellpadding="0" cellspacing="0" height="100%">
      <tr style="border-color :CadetBlue;">
        <th align="left" colspan="2">Datos del Empleado</th>
        <th nowrap style=" text-align : right;">
  			<a class=sidebtnSHW href="Javascript:Imprimir()">Imprimir</a>&nbsp;&nbsp;
  		</th>
      </tr>
      <tr style="border-color :CadetBlue;">
        <td align="left" colspan="2">&nbsp;&nbsp;&nbsp;<b>Organizaci&oacute;n: </b></td>
        <td nowrap align="left">
			<input type="Radio" name="orgfecha" value="-1" checked onclick="actFechaEstr(this.value)"> 
  			<b>a la Fecha:</b>
			<input type="Text" name="fecestr" size="10" MAXLENGTH="10" value="<%=cdate(l_fecestr)%>"  onchange="actFechaEstr(1)">
			<a href="Javascript:Ayuda_Fecha(document.datos.fecestr);actFechaEstr(1)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
			&nbsp; &nbsp;&nbsp;
			<input type="Radio" name="orgfecha" value="0" onclick="actFechaEstr(this.value)"> 
  			<b>Mostrar Hist&oacute;rico</b>
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
<script>
	Mostrar();
</script>
</body>
</html>
