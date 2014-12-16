<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'-------------------------------------------------------------------------------------------------
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'-------------------------------------------------------------------------------------------------
' Variables
' de parametros entrada
  
' de uso local  
  Dim l_evafacnro
  Dim l_evatrnro
    
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
  
' Crear registros de evaresultado para los facnro y el evldrnro
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT evaseccfactor.evaseccnro, evaseccfactor.evafacnro, evaresu.evatrnro "
  l_sql = l_sql & " FROM evaseccfactor "
  l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & " INNER JOIN evatitulo ON evatitulo.evatitnro = evafactor.evatitnro "
  l_sql = l_sql & " INNER JOIN evaresu   ON evaresu.evaseccnro  = evaseccfactor.evaseccnro AND  evaresu.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
  rsOpen l_rs, cn, l_sql, 0

  l_evafacnro = l_rs("evafacnro")
  l_evatrnro  = l_rs("evatrnro")
  
  set l_cm = Server.CreateObject("ADODB.Command")  
  do while not l_rs.eof
  		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT *  "
		l_sql = l_sql & " FROM  evaresultado "
		l_sql = l_sql & " WHERE evaresultado.evldrnro   = " & l_evldrnro
		l_sql = l_sql & " AND   evaresultado.evafacnro  = " & l_rs("evafacnro")
		rsOpen l_rs1, cn, l_sql, 0
		if l_rs1.EOF then
			l_sql = "INSERT INTO evaresultado "
			l_sql = l_sql & " (evldrnro, evafacnro, evatrnro, evaresudesc) "
			l_sql = l_sql & " VALUES (" & l_evldrnro & "," & l_rs("evafacnro")	 & ",1,'')"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		l_rs.MoveNext
		l_rs1.Close
  loop
  l_rs.Close


' MOSTRTAR evaresudes dependiendo del valor que elija como resultado -----
response.write "<script languaje='javascript'>" & vbCrLf
response.write "function Mostrar(evatrnro,evafacnro){ " & vbCrLf
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaresu.evatrnro, evaresu.evafacnro, evaresu.evaresudes "
l_sql = l_sql & " FROM evaresu "
l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
rsOpen l_rs1, cn, l_sql, 0 
l_rs1.MoveFirst
   	   
dim i 
i = 0
do while not l_rs1.eof
		response.write "if ((evatrnro == " & l_rs1(0) & ") && (evafacnro == " & l_rs1(1) & ") ) {" & vbCrLf
		response.write "document.datos.evaresudes" & l_rs1("evafacnro")& ".value = '" & l_rs1(2) & "';" & vbCrLf
		response.write "return '" & l_rs1(2) & "';" & vbCrLf
		response.write "};" & vbCrLf
l_rs1.MoveNext
loop

response.write "};" & vbCrLf
response.write "</script>" & vbCrLf

l_rs1.Close
set l_rs1 = nothing
%>

<html>
<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
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

<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<td colspan="5" align="left" class="th2">Carga de Competencias de Evaluaci&oacute;n</td>
<tr>

<tr style="border-color :CadetBlue;">
	<td>Descripci&oacute;n</td>
	<td>Resultado</td>
	<td>Observaci&oacute;n</td>
	<td>&nbsp;</td>
	<td>Observables</td>
</tr>
	
	
<%'BUSCAR evaresultados para MODIFICAR resultados ----------------------------
   Set l_rs = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evaresultado.evldrnro, evaresultado.evafacnro, "
   l_sql = l_sql & " evaresultado.evatrnro, evaresultado.evaresudesc, "
   l_sql = l_sql & " evafactor.evafacdesabr, evafactor.evafacdesext, "
   l_sql = l_sql & " evatitulo.evatitdesabr, "
   l_sql = l_sql & " evaseccfactor.orden "
   l_sql = l_sql & " FROM evaresultado "
   l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evafactor     ON evafactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evatitulo     ON evatitulo.evatitnro = evafactor.evatitnro "
   l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
   l_sql = l_sql & " AND   evaresultado.evldrnro    = " & l_evldrnro
   l_sql = l_sql & " ORDER BY evaseccfactor.orden "
   rsOpen l_rs, cn, l_sql, 0
   do while not l_rs.eof 
%>
	<tr>
		<td ><%=l_rs("evafacdesabr")%></td>	
		<td nowrap>
			<%'BUSCAR la descripcion de evaresu  ----------------------------
		    Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evaresu.evatrnro, evaresu.evaresudes, "
			l_sql = l_sql & " evatipresu.evatrvalor, evatipresu.evatrdesabr "
			l_sql = l_sql & " FROM evaresu "
			l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
			l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
			l_sql = l_sql & " AND   evaresu.evafacnro  = " & l_rs("evafacnro")
			l_sql = l_sql & " order by evatrvalor "
			rsOpen l_rs1, cn, l_sql, 0
%>
			<select name="evatrnro<%=l_rs("evafacnro")%>" onchange="Mostrar(document.datos.evatrnro<%=l_rs("evafacnro")%>.value,<%=l_rs("evafacnro")%>);">
			<option value=1> 0&nbsp;&nbsp; Sin Evaluar</option>
			<%l_rs1.MoveFirst
			  do while not l_rs1.eof%>
				<option value=<%=l_rs1("evatrnro")%>><%=l_rs1("evatrvalor")%>&nbsp;&nbsp;&nbsp;<%=l_rs1("evatrdesabr")%></option>
			<%l_rs1.MoveNext
			loop 
			l_rs1.Close
			set l_rs1 = nothing%>
			</select>
			<input  disabled type=text name="evaresudes<%=l_rs("evafacnro")%>">
			<script>document.datos.evatrnro<%=l_rs("evafacnro")%>.value='<%=l_rs("evatrnro")%>'</script>
			<script>Mostrar(document.datos.evatrnro<%=l_rs("evafacnro")%>.value,<%=l_rs("evafacnro")%>);</script>
			</td>
		<td>
<%
'response.write(l_rs("evldrnro"))
'response.write("evaresdesc=")
'response.write(l_rs("evaresdesc"))
'response.write("evafacnro=")
'response.write(l_rs("evafacnro"))%>
			<textarea name="evaresudesc<%=l_rs("evafacnro")%>" cols=20 rows=1><%=trim(l_rs("evaresudesc"))%></textarea>
		</td>
		<td nowrap valign=top><a href=# onclick="grabar.location='grabar_resultados_evaluacion_00.asp?evafacnro=<%=l_rs("evafacnro")%>&evldrnro=<%=l_evldrnro%>&evaresudesc='+escape(document.datos.evaresudesc<%=l_rs("evafacnro")%>.value)+'&evatrnro='+document.datos.evatrnro<%=l_rs("evafacnro")%>.value;document.datos.grabado<%=l_rs("evafacnro")%>.value='G';">Grabar</a>
			<input type="text" readonly disabled name="grabado<%=l_rs("evafacnro")%>" size="1">
			</td>
		
		<%  Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evadescomp.evadcdes "
			l_sql = l_sql & " FROM evadescomp "
			l_sql = l_sql & " WHERE evadescomp.evafacnro = " & l_rs("evafacnro")
			rsOpen l_rs1, cn, l_sql, 0
			if not l_rs1.eof then%>
				<td valign=top><a href=# onclick="alert('<%=l_rs1("evadcdes")%>')">?</a></td>
			<%else%>	
				<td valign=top>?</td>
			<%end if%>	
		</tr>
		<%
		l_rs1.Close
		set l_rs1 = nothing
		
		l_rs.Movenext
		loop
		l_rs.Close%>

</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
