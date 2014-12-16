<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!-----------------------------------------------------------------------------------------------
Archivo		: consulta_no_busquedas_eyp_00.asp
Descripción	: Consulta de las busquedas en las que no aparece un candidato.
Autor 		: Lisandro Moro
Fecha		: 27-05-04
Modificado	:
				30/03/2005 - Martin Ferraro - Se cambio los len por la funcion BD_longitud para Oracle

-------------------------------------------------------------------------------------------------
-->
<%
on error goto 0
 Dim l_rs
 Dim l_sql
 
 Dim l_apenom
 Dim l_ant_nrodoc
 Dim l_sig_nrodoc
 Dim l_ant_ternro
 Dim l_sig_ternro
 
 Dim l_ternro
 Dim l_nrodoc
 Dim l_estado
 
' l_ternro	= request.QueryString("ternro")
' l_nrodoc	= request.QueryString("nrodoc")
' l_estado	= request.QueryString("estado")
 l_estado	= -1
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
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
 
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>B&uacute;squedas - Empleos y Postulantes - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>

function Incorporar(){
	if (document.ifrm.datos.cabnro.value == ''){
		alert('Debe seleccionar una busqueda.');
	}else{
		abrirVentanaH("consulta_no_busquedas_eyp_02.asp?ternro=<%= l_ternro %>&reqbusnro=" + document.ifrm.document.datos.cabnro.value,'',1000,1000);
	}
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr style="border-color :CadetBlue;">
    	<th align="left">B&uacute;squedas a las que no está asociado el Empleado</th>
        <th nowrap align="right">
			<a class=sidebtnABM href="Javascript:Incorporar();">Incorporar</a>
		</th>
	</tr>

	<tr>
		<td colspan="2" height="5">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" height="10">&nbsp;</td>
	</tr>
	
	<tr valign="top" height="50%">
    	<td colspan="2" align="center">
		 <table style="width:97%; border-color:gray; border-width: 1; border-style:solid;" cellspacing="2" height="100%">
			<tr>
			   	<td width="100%"align="center" colspan="2">
					<b>B&uacute;squedas</b>
			    </td>
			</tr>
			<tr>
		    	<td width="100%" align="center" height="100%" colspan="2">
					<iframe name="ifrm" scrolling="Yes" src="consulta_no_busquedas_eyp_01.asp?ternro=<%= l_ternro %>" width="98%" height="98%"></iframe> 
			    </td>
			</tr>
		 </table>
	    </td>
	</tr>
	
	<tr>
		<td colspan="2" height="10">&nbsp;</td>
	</tr>
	
	<tr valign="top" height="50%">
    	<td colspan="2" align="center">
		 <table style="width:97%; border-color:gray; border-width: 1; border-style:solid;" cellspacing="2" height="100%">
		  <tr>
	    	<td width="100%" align="center">
				<b>Condición</b>
			</td>
		  </tr>
		  <tr>
	    	<td width="100%" align="center" height="100%" colspan="2">
				<textarea name="textsql" readonly style="width:100%;height:100%;" ></textarea>
			</td>
		  </tr>
		 </table>
	    </td>
	</tr>
	
	<tr>
		<td colspan="2" height="10"></td>
	</tr>
</table>
</body>
</html>