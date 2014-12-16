<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'--------------------------------------------------------------------------
'Archivo       : carga_cardinales_eva_00.asp
'Descripcion   : visualizacion de competencias Cardinales
'Creacion      : 05-05-2004
'Autor         : CCRossi
'Modificacion  : 11-05-2004 CCRossi agregar busqueda de observables por tipoestructura
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'				24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'--------------------------------------------------------------------------

' Variables

' de parametros de entrada---------------------------------------
  Dim l_evaseccnro
  Dim l_evldrnro
  
' de uso local  
  Dim l_evafacnro
  dim l_evatitdesabr
  
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

  dim l_empleado 
  dim l_estrnro    
  dim l_gerencia  
  
' parametros de entrada---------------------------------------  
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro")

  'response.write("<script> l_evatitdesabr
' Crear registros de evaresultado para los facnro y el evldrnro
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT evaseccfactor.evaseccnro, evaseccfactor.evafacnro, evatitulo.evatitdesabr "
  l_sql = l_sql & " FROM evaseccfactor "
  l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & " INNER JOIN evatitulo ON evatitulo.evatitnro = evafactor.evatitnro "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.EOF then
	l_evafacnro = l_rs("evafacnro")
	l_evatitdesabr = l_rs("evatitdesabr")
 end if  	
 l_rs.close
 set l_rs=nothing

l_empleado = ""
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empleado  "
l_sql = l_sql & " FROM evacab"
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro"
l_sql = l_sql & " WHERE evldrnro=" & l_evldrnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then	
	l_empleado= l_rs("empleado")
else
	l_empleado = ""
end if	
l_rs.Close
set l_rs=nothing

 
 if trim(l_empleado)<>"" then 
'buscar la gerencia -----------------------------------------------------------------
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT estrdabr, htetdesde, estructura.estrnro  "
l_sql = l_sql & " FROM his_estructura "
l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
l_sql = l_sql & " WHERE his_estructura.ternro=" & l_empleado
l_sql = l_sql & " AND   his_estructura.tenro = 6 " 
l_sql = l_sql & " AND   his_estructura.htethasta IS NULL " 
l_sql = l_sql & " ORDER BY his_estructura.htetdesde DESC " 
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then	
	l_gerencia = l_rs("estrdabr")
	l_estrnro  = l_rs("estrnro")
else
	l_gerencia = "--"
	l_estrnro  = ""
end if	
l_rs.Close
set l_rs=nothing

end if
'response.write("<script>alert('"& l_evatitdesabr &"');</script>")
%>

<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Carga de Competencias de Evaluaci&oacute;n - Evaluaci&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" >
<form name="datos">

<table border="0" cellpadding="0" cellspacing="0" height="90%" width="100%">
<tr style="border-color :CadetBlue;">
	<td colspan="5" align="left" class="th2"><%=l_evatitdesabr %></td>
<tr>

<tr style="border-color :CadetBlue;">
	<td><b>Descripci&oacute;n</b></td>
	<td><b>Ponderaci&oacute;n por Gerencia</b></td>
</tr>
<%'BUSCAR factores ----------------------------
    Set l_rs = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evafactor.evafacnro, "
   'l_sql = l_sql & " evaresultado.evldrnro, "
   l_sql = l_sql & " evafactor.evafacdesabr, evafactor.evafacdesext, "
   l_sql = l_sql & " evatitulo.evatitdesabr, "
   l_sql = l_sql & " evaseccfactor.orden "
   'l_sql = l_sql & " FROM evaresultado "
   l_sql = l_sql & " FROM evaseccfactor "
   'l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evafactor     ON evafactor.evafacnro = evaseccfactor.evafacnro "
   l_sql = l_sql & " INNER JOIN evatitulo     ON evatitulo.evatitnro = evafactor.evatitnro "
   l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
   'l_sql = l_sql & " AND   evaresultado.evldrnro    = " & l_evldrnro
   l_sql = l_sql & " ORDER BY evaseccfactor.orden "
   rsOpen l_rs, cn, l_sql, 0
   
   do while not l_rs.eof 
%>

<tr>
	<td valign=middle><b><%=l_rs("evafacdesabr")%></b></td>	
	<%
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evadescomp.evadcdes "
	l_sql = l_sql & " FROM evadescomp "
	l_sql = l_sql & " WHERE evadescomp.evafacnro = " & l_rs("evafacnro")
	if trim(l_estrnro) <>"" then
    l_sql = l_sql & " AND evadescomp.estrnro = " & l_estrnro
	end if
	rsOpen l_rs1, cn, l_sql, 0
	if not l_rs1.eof then%>
		<td valign=middle><%=l_rs1("evadcdes")%></td>
	<%else%>	
		<td valign=middle><a href="#" onclick="alert('No hay definida Ponderación para la Gerencia y la Competencia.');">?</a></td>
	<%end if

	  l_rs1.Close
	  set l_rs1 = nothing%>	
</tr>
<%
l_rs.MoveNext
loop
l_rs.close
set l_rs=nothing

%>
</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
