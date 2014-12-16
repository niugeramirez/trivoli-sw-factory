<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'------------------------------------------------------------------------------------
' Nombre		: ver_responsabilidad_eva_00.asp
' Descripcion	: permite ver las notas, para una dada seccion 
' Autor			: Leticia Amadio
' Fecha 		: 28-12-2004    
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 21-08-2007 - Diego Rosso - Se agrego src="blanc.asp" para https
'-------------------------------------------------------------------------------------

on error goto 0

Dim l_descripcion 
Dim l_evaengdesabr
    
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_rs2
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evaseccnro
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro")

Set l_rs  = Server.CreateObject("ADODB.RecordSet") 
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
set l_cm  = Server.CreateObject("ADODB.Command")	
  
  
 'HARCODED---------------------- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'  l_evaseccnro = 4
'  l_evldrnro   = 1
  
' Crear registros de evaNOTAS para evldrnro y el tipo nota
l_sql = "SELECT evatnnro "
l_sql = l_sql & "FROM evaseccnota "
l_sql = l_sql & "WHERE evaseccnota.evaseccnro = " & l_evaseccnro
rsOpen l_rs, cn, l_sql, 0

l_sql = "SELECT evaengdesabr  "
l_sql = l_sql & "FROM  evaengage "
l_sql = l_sql & " INNER JOIN evaproyecto  ON evaproyecto.evaengnro=evaengage.evaengnro"
l_sql = l_sql & " INNER JOIN evaevento    ON evaevento.evaproynro=evaproyecto.evaproynro"
l_sql = l_sql & " INNER JOIN evacab       ON evacab.evaevenro=evaevento.evaevenro"
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro"
l_sql = l_sql & " WHERE evadetevldor.evldrnro   = " & l_evldrnro
rsOpen l_rs1, cn, l_sql, 0

if not l_rs1.EOF then
	l_evaengdesabr=l_rs1("evaengdesabr")
else
	l_evaengdesabr=""
end if
	
l_rs1.Close


do while not l_rs.eof
	
	l_sql = "SELECT *  "
	l_sql = l_sql & "FROM  evanotas "
	l_sql = l_sql & "WHERE evanotas.evldrnro   = " & l_evldrnro
	l_sql = l_sql & "AND   evanotas.evatnnro  = " & l_rs("evatnnro")
	rsOpen l_rs1, cn, l_sql, 0
	
	if l_rs1.EOF then
		
		l_sql = " SELECT evatndesabr  "
		l_sql = l_sql & " FROM evatiponota "
		l_sql = l_sql & " WHERE evatnnro  = " & l_rs("evatnnro")
		rsOpen l_rs2, cn, l_sql, 0
		
		if  (INSTR(ucase(l_rs2("evatndesabr")),"ENGAGE") > 0) then
			l_descripcion = l_evaengdesabr   ' descripc del engagement
		else
			l_descripcion = " "
		end if 
		l_rs2.Close
	
		l_sql = "INSERT INTO evanotas "
		l_sql = l_sql & "(evldrnro, evatnnro, evanotadesc) "
		l_sql = l_sql & " VALUES (" & l_evldrnro & "," &  l_rs("evatnnro")
		l_sql = l_sql &  ",'"& l_descripcion &"')"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if
	l_rs.MoveNext
	l_rs1.Close
loop
l_rs.Close
 
%>

<html>
<head>
<link href="../<%= c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Notas de Evaluaci&oacute;n - Evaluaci&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos">
<input type="Hidden" name="terminarsecc" value="SI">

<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<th colspan="5" align="left" class="th2">Notas de Evaluaci&oacute;n</th>
<tr>
<tr style="border-color :CadetBlue;">
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr style="border-color :CadetBlue;">
	<td><strong> Tipo de Nota</strong></td>
	<td><strong>Nota</strong></td>
	<td>&nbsp;</td>
</tr>
<tr style="border-color :CadetBlue;">
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
	
<%'BUSCAR evaNotas para MODIFICAr resultados ------------------------------
   Set l_rs = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evldrnro, evanotas.evatnnro, evanotadesc, evatndesabr, evatndesext,evaseccnota.orden "
   l_sql = l_sql & "FROM evanotas "
   l_sql = l_sql & "INNER JOIN evaseccnota ON evaseccnota.evatnnro = evanotas.evatnnro "
   l_sql = l_sql & "INNER JOIN evatiponota ON evatiponota.evatnnro = evanotas.evatnnro "
   l_sql = l_sql & "WHERE evaseccnota.evaseccnro = " & l_evaseccnro
   l_sql = l_sql & " AND   evanotas.evldrnro      = " & l_evldrnro
   l_sql = l_sql & " ORDER BY evaseccnota.orden "
   rsOpen l_rs, cn, l_sql, 0
   do while not l_rs.eof %>
   <tr>
		<td valign=top><%=l_rs("evatndesabr")%></td>
		<td>
		<textarea readonly disabled name="evanotadesc<%=l_rs("evatnnro")%>"  maxlength=255 size=255 cols=40 rows=6><%=trim(l_rs("evanotadesc"))%></textarea>
		</td>
		<td>&nbsp;</td>
    </tr>
  <%l_rs.Movenext
  loop
  l_rs.Close%>

</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
