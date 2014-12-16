<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'---------------------------------------------------------------------------------------------------------
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'---------------------------------------------------------------------------------------------------------
' Variables
' de parametros entrada
  
' de uso local  
    
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evaseccnro
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro")

  
 'HARCODED---------------------- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'  l_evaseccnro = 4
'  l_evldrnro   = 1
  
' Crear registros de evaNOTAS para evldrnro y el tipo nota
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT evatnnro "
  l_sql = l_sql & "FROM evaseccnota "
  l_sql = l_sql & "WHERE evaseccnota.evaseccnro = " & l_evaseccnro
  rsOpen l_rs, cn, l_sql, 0

  set l_cm = Server.CreateObject("ADODB.Command")  
  do while not l_rs.eof
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT *  "
		l_sql = l_sql & "FROM  evanotas "
		l_sql = l_sql & "WHERE evanotas.evldrnro   = " & l_evldrnro
		l_sql = l_sql & "AND   evanotas.evatnnro  = " & l_rs("evatnnro")
		rsOpen l_rs1, cn, l_sql, 0
		if l_rs1.EOF then
			l_sql = "INSERT INTO evanotas "
			l_sql = l_sql & "(evldrnro, evatnnro, evanotadesc) "
			l_sql = l_sql & " VALUES (" & l_evldrnro & "," &  l_rs("evatnnro")
			l_sql = l_sql &  ",'')"
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
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Carga de Notas de Evaluaci&oacute;n - Evaluaci&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos">

<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<td colspan="5" align="left" class="th2"><%if ccodelco=-1 then%>Conclusiones<%else%>Carga de Notas de Evaluaci&oacute;n<%end if%></td>
<tr>
<tr style="border-color :CadetBlue;">
	<td>Tipo de Nota</td>
	<td>Nota</td>
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
		<td valign=top><%=l_rs("evatndesext")%></td>
		<td>
		<textarea name="evanotadesc<%=l_rs("evatnnro")%>"  maxlength=255 size=255 cols=40 rows=6><%=trim(l_rs("evanotadesc"))%></textarea>
		</td>
		<td valign=top><a href=# onclick="grabar.location='grabar_notas_evaluacion_00.asp?evatnnro=<%=l_rs("evatnnro")%>&evldrnro=<%=l_evldrnro%>&evanotadesc='+document.datos.evanotadesc<%=l_rs("evatnnro")%>.value;document.datos.grabado<%=l_rs("evatnnro")%>.value='G';">Grabar</a>
			<input type="text" readonly disabled name="grabado<%=l_rs("evatnnro")%>" size="1">
		</td>
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
