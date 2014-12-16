<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'=====================================================================================
'Archivo  : carga_general_eva_00.asp
'Objetivo : ABM de objetivos de evaluacion
'Fecha	  : 17-05-2004
'Autor	  : CCRossi
'Modificado : 13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'				24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
 
' Variables
' de uso local  
  Dim l_existe  
  Dim l_visfecha
  Dim l_visdesc
  dim l_evacabnro
  dim l_lista  
  dim l_caracteristica  
  dim l_nombre
  dim l_puntaje
      
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evldrnro
  Dim l_evaseccnro
  
' parametros de entrada---------------------------------------  
  l_evldrnro   = Request.QueryString("evldrnro")
  l_evaseccnro = Request.QueryString("evaseccnro")

'___________________________________________________________________________________
function PasarComaAPunto(valor)
	dim l_numero
	dim l_ubicacion
	dim l_entero
	dim l_decimal
	l_numero = trim(valor)
	l_ubicacion = InStr(l_numero, ",")
	if l_ubicacion > 1 then
		l_ubicacion = l_ubicacion  - 1
		l_entero = left(l_numero, l_ubicacion)
		l_ubicacion = l_ubicacion  + 1
		l_decimal = right(l_numero, (len(l_numero) - l_ubicacion))
    	l_numero = l_entero & "." & l_decimal
    	PasarComaAPunto = l_numero
    else
		PasarComaAPunto = valor
	end if
end function	
'____________________________B O D Y _______________________________________________

'buscar la evacab
 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evacabnro  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs1, cn, l_sql, 0
 if not l_rs1.EOF then
	l_evacabnro = l_rs1("evacabnro")
 end if
 l_rs1.close
 set l_rs1=nothing
 
'Crear registros de evaVistos para todos los evldrnro ---------------------------

 Set l_rs = Server.CreateObject("ADODB.RecordSet")	
 l_sql = "SELECT DISTINCT empleado.ternro, empleado.empleg, empleado.terape, empleado.ternom, habilitado, "
 l_sql = l_sql & " evatipevalua.evatevdesabr, evadetevldor.evldrnro, "
 l_sql = l_sql & " evaoblieva.evaobliorden, evacab.cabaprobada "
 l_sql = l_sql & " FROM evacab "
 l_sql = l_sql & " inner join evadetevldor on evacab.evacabnro= evadetevldor.evacabnro "
 l_sql = l_sql & " inner join evaoblieva on evaoblieva.evatevnro= evadetevldor.evatevnro "
 l_sql = l_sql & " left join empleado on evadetevldor.evaluador= empleado.ternro "
 l_sql = l_sql & " inner join evatipevalua on evadetevldor.evatevnro= evatipevalua.evatevnro "
 l_sql = l_sql & " inner join evasecc on evadetevldor.evaseccnro= evasecc.evaseccnro "
 l_sql = l_sql & " WHERE evacab.evacabnro = " & l_evacabnro
 l_sql = l_sql & " AND evadetevldor.evaseccnro=" & l_evaseccnro
 l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden " 
'response.write l_sql & "<br>"
 rsOpen l_rs, cn, l_sql, 0 
 l_lista="0"
 do until l_rs.eof
   l_lista = l_lista & "," & l_rs("evldrnro")
   
   Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT *  "
	l_sql = l_sql & " FROM  evavistos "
	l_sql = l_sql & " WHERE evavistos.evldrnro   = " & l_rs("evldrnro")
	rsOpen l_rs1, cn, l_sql, 0
	'response.write(l_sql)
	if l_rs1.EOF then
		l_visfecha = cambiafecha(Date(),"","")
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "insert into evavistos "
		l_sql = l_sql & "(evldrnro,visfecha) "
		l_sql = l_sql & "values (" & l_rs("evldrnro") &","&l_visfecha &")"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if
	l_rs1.Close
	set l_rs1=nothing
	
   l_rs.MoveNext
 loop 
 l_rs.Close
 set l_rs=nothing 
%>

<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Evaluaci&oacute;n General y Comentarios - Gesti&oacute;n del Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function aumentar(anterior, nuevo)
{
	
	if (Number(nuevo.value)< Number(anterior.value))
		nuevo.value=Number(nuevo.value) + 0.5;
	else	
	if (Number(anterior.value)< Number(5))
		nuevo.value=Number(anterior.value) + 0.5;
	else
	{
		alert('La Puntuación no puede superar a 5.')
		nuevo.focus();
	}
	
}
function disminuir(anterior, nuevo)
{
	if (Number(nuevo.value)>Number(anterior.value))
		nuevo.value=Number(nuevo.value) - 0.5;
	else	
	if (Number(anterior.value)> Number(0))
		nuevo.value=Number(anterior.value) - 0.5;
	else
	{
		alert('La Puntuación no puede ser inferior a 0.')
		nuevo.focus();
	}
	
}
</script>
<style>
.rev
{
	font-size: 10;
	border-style: none;
}
</style>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos">

<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<td colspan="6" align="left" class="th2">Evaluaci&oacute;n General y Comentarios</td>
<tr>
<tr style="border-color :CadetBlue;">
	<td><b>Evaluador</b></td>
	<td><b>Comentario</b></td>
	<td><b>Puntuaci&oacute;n Obtenida</b></td>
	<td><b>Evaluaci&oacute;n General</b></td>
	<td><b>Fecha</b></td>
	<td>&nbsp;</td>
</tr>
	
<%	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evavistos.evldrnro, visdesc,visfecha, evatevdesabr, puntaje, puntajemanual "
	l_sql = l_sql & " FROM  evavistos "
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro=evavistos.evldrnro"
	l_sql = l_sql & " LEFT  JOIN evapuntaje   ON evapuntaje.evacabnro=evadetevldor.evacabnro "
	l_sql = l_sql & "			AND evapuntaje.evatevnro=evadetevldor.evatevnro "
	l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro = evadetevldor.evatevnro"
	l_sql = l_sql & " WHERE evavistos.evldrnro IN (" & l_lista & ")"
	rsOpen l_rs1, cn, l_sql, 0
	do while not l_rs1.eof
		if Int(l_evldrnro) <> l_rs1("evldrnro") then
			l_caracteristica = "readonly disabled"
			l_nombre = l_evldrnro
		else	
			l_caracteristica = ""
			l_nombre = ""
		end if
		
		if trim(l_rs1("puntajemanual"))<>"" then
			l_puntaje= l_rs1("puntajemanual")
		else
			l_puntaje= l_rs1("puntaje")
		end if%>
	<tr>
		<td>
			<b><%=l_rs1("evatevdesabr")%></b>
		</td>
		<td>
			<textarea <%=l_caracteristica%> name="visdesc<%=l_nombre%>"  maxlength=200 size=200 cols=40 rows=5><%=trim(l_rs1("visdesc"))%></textarea>
		</td>
		<td align=center>
			<input <%=l_caracteristica%> readonly style="background : #e0e0de;" class="rev" type="text" name="obtenida<%=l_nombre%>" size="2" maxlength="2" value="<%=PasarComaAPunto(l_rs1("puntaje"))%>">
		</td>
		<td align=center>
			<input type="hidden" name="puntajeant<%=l_nombre%>" size="5" maxlength="4" value="<%=PasarComaAPunto(l_rs1("puntaje"))%>">
			<input <%=l_caracteristica%> readonly style="background : #e0e0de;" class="rev" type="text" name="puntaje<%=l_nombre%>" size="2" maxlength="2" value="<%=PasarComaAPunto(l_puntaje)%>">
			<br>
			<a <%=l_caracteristica%> href="Javascript:aumentar(document.datos.puntajeant<%=l_nombre%>,document.datos.puntaje<%=l_nombre%>);"><font type="tahoma" size=1>Aumentar</font></a>&nbsp;
			<a <%=l_caracteristica%> href="Javascript:disminuir(document.datos.puntajeant<%=l_nombre%>,document.datos.puntaje<%=l_nombre%>);"><font type="tahoma" size=1>Bajar</font></a>
		</td>
		<td>
			<input <%=l_caracteristica%> type="text" name="visfecha<%=l_nombre%>" size="10" maxlength="10" value="<%=l_rs1("visfecha")%>">
			<%if trim(l_caracteristica)="" then %>
			<a href="Javascript:Ayuda_Fecha(document.datos.visfecha<%=l_nombre%>)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
			<%end if%>
		</td>
		<td valign=middle>
			<%if trim(l_caracteristica)="" then %>
			<a href=# onclick="if (validarfecha(document.datos.visfecha<%=l_nombre%>)) {grabar.location='grabar_vistos_evaluacion_00.asp?tipo=M&evldrnro=<%=l_evldrnro%>&visdesc='+escape(document.datos.visdesc.value)+'&visfecha='+document.datos.visfecha.value+'&puntaje='+document.datos.puntaje.value;document.datos.grabado.value='M';}">Actualizar</a>			
			<input type="text" readonly disabled name="grabado" size="1">
			<%end if%>
		</td>
    </tr>
    <%l_rs1.MoveNext
   loop
   l_rs1.close
   set l_rs1=nothing
%>
</form>	
</table>
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
<!--iframe name="grabar" style="width:500;height:100"-->
</iframe>

</body>
</html>
