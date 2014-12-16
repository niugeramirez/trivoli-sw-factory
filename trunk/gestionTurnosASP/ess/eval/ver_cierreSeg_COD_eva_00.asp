<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'=====================================================================================
'Archivo  : ver_cierreSeg_COD_eva_00.asp
'Objetivo : Cierre de una etapa (1: planificacion, 2: seguimiento, 3 evaluacion)
'Fecha	  : 24-02-2005
'Autor	  : CCRossi
'Modificacion: 
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================

' Variables
' de uso local  
  Dim l_existe  
  Dim l_evareunion
  Dim l_evafecha
  Dim l_evaobser
  Dim l_evaetapa

  Dim l_texto
    
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evldrnro = Request.QueryString("evldrnro")
  l_evaetapa = 2
  
' Crear registros de evafirm para evldrnro y el tipo nota
   Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT *  "
   l_sql = l_sql & " FROM  evacierre "
   l_sql = l_sql & " WHERE evacierre.evldrnro =  " & l_evldrnro
   l_sql = l_sql & "   AND evacierre.evaetapa = " & l_evaetapa
   rsOpen l_rs1, cn, l_sql, 0
   if l_rs1.EOF then
		l_texto= "No hay datos cargados."
   else
  		l_evareunion=l_rs1("evareunion")
  		l_evaobser= l_rs1("evaobser")
   end if
   l_rs1.Close
   set l_rs1=nothing
   
   Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evaevefecha  "
   l_sql = l_sql & " FROM  evaevento "
   l_sql = l_sql & " INNER JOIN evacab ON evacab.evaevenro=evaevento.evaevenro "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro "
   l_sql = l_sql & " WHERE evadetevldor.evldrnro =  " & l_evldrnro
   rsOpen l_rs1, cn, l_sql, 0
   if not l_rs1.eof then
		l_evafecha = l_rs1("evaevefecha")
   end if
   l_rs1.Close
   set l_rs1=nothing
%>

<html>
<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Cierre del Seguimiento - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
</script>

<style>
.rev
{
	font-size: 11;
	border-style: none;
}
</style>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="<%if trim(l_texto)="" then%><% if l_evareunion=0 then%>document.datos.evareunion.value=0<%else%>document.datos.evareunion.value=-1<%end if%><%end if%>;">
<form name="datos">

<table border="0" cellpadding="0" cellspacing="0">
<%if trim(l_texto)<>"" then%>	
<tr>
 	<td colspan=2><%=l_texto%></td>
</tr>
<%else%>
	
<tr>
 	<td align=right>
 		<b>¿Se realiz&oacute; la reuni&oacute;n de Seguimiento?</b>
 	</td>
 	<td align=left>
 		<input type="hidden" name="evareunion" value="<%=l_evareunion%>">
 		<%If l_evareunion = 0 then%>NO<%else%>SI<%End If%>
 	</td>
</tr>
<tr>		
	<td align=right>
		<b>Observaci&oacute;n:</b>
	</td>
	<td align=left>
		<textarea readonly style='background : #e0e0de;' class="rev" name="evaobser"  maxlength=200 size=200 cols=40 rows=5><%=trim(l_evaobser)%></textarea>
	</td>
</tr>
<tr>
	<td align=right>
		<b>Fecha de Evaluaci&oacute;n:</b>
	</td>
	<td align=left>
		<input readonly style="background : #e0e0de;" class="rev" type="text" name="evafecha" size="10" maxlength="10" value="<%=l_evafecha%>">
	</td>
</tr>
<tr>
	<td valign=top align=right colspan=2>
		<input class="rev" type="hidden" style="background : #e0e0de;" readonly disabled name="grabado" size="1">
	</td>
</tr>
</form>	
<%end if
cn.Close
set cn = Nothing
%>
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
