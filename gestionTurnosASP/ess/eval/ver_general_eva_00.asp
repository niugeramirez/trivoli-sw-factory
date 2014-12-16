<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'=====================================================================================
'Archivo  : ver_general_eva_00.asp
'Objetivo : ABM de objetivos de evaluacion
'Fecha	  : 17-05-2004
'Autor	  : CCRossi
'Modificación : CCRossi - 02-11-2004 Agregar tablita de referencia abajo
'				Leticia Amadio - 26-01-2005 
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
 on error  goto 0
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
  dim l_puntajemanual
  dim l_evatevnro
  dim l_verevaluacion


  dim l_evaluador ' guarda el empleg del evaluador del evadetevldor, para comparar con el logeado.
  dim l_mostrar '1 o 0 si teine que mostrsr la observacion. 
 
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evldrnro
  Dim l_evaseccnro
  Dim l_empleg

 l_empleg = Session("empleg")
 if trim(l_empleg)="" then
	l_empleg = Request.QueryString("empleg")
 end if	
'response.write("<script>alert('"& l_empleg&"')</script>")

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


' _______________________________________________________________________________
'buscar la evacab y evatevnro 
 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evacabnro,empleado.empleg, evatevnro  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs1, cn, l_sql, 0
 if not l_rs1.EOF then
	l_evacabnro = l_rs1("evacabnro")
	l_evaluador = l_rs1("empleg")
	l_evatevnro = l_rs1("evatevnro")
 end if
 l_rs1.close
 set l_rs1=nothing

'  XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXx evatevnro es el del evldrnro y no de quien se loguea!!!
'  ver si el evaluador termino la sección  ________________________________
'response.write l_evldrnro & " --"
'response.write l_evatevnro
'response.write " -  " & cautoevaluador

' if l_evatevnro = cautoevaluador then 
	  
	 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	 l_sql = " SELECT evldorcargada "
	 l_sql = l_sql & " FROM evadetevldor "
	 l_sql = l_sql & " WHERE evaseccnro = " & l_evaseccnro & " AND evatevnro = " & cevaluador
	 rsOpen l_rs1, cn, l_sql, 0
  	  
	 l_verevaluacion = 0
 	  
	 if not l_rs1.EOF then
 		if l_rs1("evldorcargada")= -1 then
 			l_verevaluacion = -1
		end if
	 end if
 	  
	 l_rs1.close
	 set l_rs1=nothing
'else 
	'l_verevaluacion = -1
'end if

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
		'l_visfecha = cambiafecha(Date(),"","")
		'l_visfecha = null
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "insert into evavistos "
		l_sql = l_sql & "(evldrnro) "
		l_sql = l_sql & "values (" & l_rs("evldrnro") &")"
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

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT puntaje, puntajemanual "
l_sql = l_sql & " FROM  evacab "
l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then
	l_puntaje= l_rs("puntaje")
	l_puntajemanual= l_rs("puntajemanual")
end if
l_rs.close
set l_rs=nothing
	
if trim(l_puntajemanual)="" or isnull(l_puntajemanual) then
	l_puntajemanual= l_puntaje
else
	if cdbl(l_puntajemanual)=0 then	
	l_puntajemanual= l_puntaje
	end if
end if

%>

<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Evaluaci&oacute;n General y Comentarios - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
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

</script>
<style>
.rev
{
	font-size: 10;
	border-style: none;
}
.blanc
{
	font-size: 10;
	border-style: none;
	background : transparent;
}
</style>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos">

<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<th colspan="4" align="left" class="th2">Evaluaci&oacute;n General y Comentarios</th>
<tr>
<tr>
	<td>&nbsp;</td>
	<td align=center valign=top><b>Evaluaci&oacute;n General</b>
		<% if l_verevaluacion=-1 then%>
			&nbsp;
			<input readonly style="background : #e0e0de;" class="rev" type="text" name="puntajemanual" size="2" maxlength="2" value="<%=PasarComaAPunto(l_puntajemanual)%>">
		 <%else%>
			<input readonly style="background : #e0e0de;" class="rev" type="text" name="puntajemanual" size="10" maxlength="2" value="No habilitado">
		 <%end if%>
	</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr style="border-color :CadetBlue;">
	<td><b>&nbsp;</b></td>
	<td><b>Comentario</b></td>
	<td><b>Fecha</b></td>
	<td>&nbsp;</td>
</tr>

				
<%	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evavistos.evldrnro, visdesc,visfecha, evatevdesabr,empleado.empleg   "
	l_sql = l_sql & " FROM  evavistos "
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro=evavistos.evldrnro"
	l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
	l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro = evadetevldor.evatevnro"
	l_sql = l_sql & " WHERE evavistos.evldrnro IN (" & l_lista & ")"
	rsOpen l_rs1, cn, l_sql, 0
	do while not l_rs1.eof
		l_caracteristica = "readonly disabled"
		
		l_nombre    = l_evldrnro
		
		l_evaluador = l_rs1("empleg")
		' si el evaluador actual no es el usuario logeado no mostar observaciones!
	    l_mostrar="0"
		if trim(l_empleg)<>"" and not isNull(l_empleg) then
			if trim(l_empleg) = trim(l_evaluador) then
				l_mostrar = "1"
			else	
				l_mostrar = "0"
			end if
		else
			l_mostrar="1"
		end if
%>
	<tr>
		<td>
			<b><%=l_rs1("evatevdesabr")%></b>
		</td>
		<td>
			<textarea <%=l_caracteristica%> name="visdesc<%=l_nombre%>"  maxlength=200 size=200 cols=50 rows=4><% if l_verevaluacion then %> <%=trim(l_rs1("visdesc"))%> <%else%> No habilitado <% end if%></textarea>
		</td>
		<td>
			<b>Firmada el </b>
			<input <%=l_caracteristica%> type="text" name="visfecha<%=l_nombre%>" size="10" maxlength="10" value="<%=l_rs1("visfecha")%>">
			<%if trim(l_caracteristica)="" then %>
			<a href="Javascript:Ayuda_Fecha(document.datos.visfecha<%=l_nombre%>)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
			<%end if%>
		</td>
		<td valign=top>
			<%if trim(l_caracteristica)="" then %>
			<a href=# onclick="if (validarfecha(document.datos.visfecha<%=l_nombre%>)) {grabar.location='grabar_vistos_evaluacion_00.asp?tipo=M&evldrnro=<%=l_evldrnro%>&visdesc='+escape(document.datos.visdesc.value)+'&visfecha='+document.datos.visfecha.value;document.datos.grabado.value='M';}">Actualizar</a>			
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
	<tr>
		<td colspan=4>&nbsp;
		</td>
    </tr>
	<tr>
		<td width="30%">&nbsp;</td>
		<td colspan=2 align=center colspan=4>
			<table border=1 >
			<tr>
				<td><b>Puntuaci&oacute;n</b></td>
				<td><b>Correspondencia Literal</b></td>
			</tr>
			<tr>
				<td>0&nbsp;1</td>
				<td>Significativamente debajo de los requerimientos</td>
			</tr>
				<td>1,5&nbsp;;&nbsp;2&nbsp;y&nbsp;2,5</td>
				<td>Debajo de los requerimientos</td>
			<tr>
				<td>3</b></td>
				<td>Cumple con los requerimientos</td>
			</tr>
			<tr>
				<td>3,5&nbsp;y&nbsp;4</b></td>
				<td>Encima de los requerimientos</td>
			</tr>
			<tr>
				<td>4,5&nbsp;y&nbsp;5</b></td>
				<td>Supera significativamente los requerimientos</td>
			</tr>
			</table>
		</td>
		<td width="10%">&nbsp;</td>
    </tr>
</table>
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
