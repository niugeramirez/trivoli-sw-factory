<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'************************************************************************************************************
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'************************************************************************************************************
' Variables
' de uso local  
  Dim l_existe  
  Dim l_visfecha
  Dim l_visdesc
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evldrnro = Request.QueryString("evldrnro")
  
' Crear registros de evafirm para evldrnro y el tipo nota
   Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT *  "
   l_sql = l_sql & " FROM  evavistos "
   l_sql = l_sql & " WHERE evavistos.evldrnro   = " & l_evldrnro
   rsOpen l_rs1, cn, l_sql, 0

'   response.write(l_sql)
   if l_rs1.EOF then
   
    l_existe = "no"
  	l_visfecha = Date()
	l_visdesc  = ""
   else
  	l_existe = "si"
  	l_visfecha = l_rs1("visfecha")
	l_visdesc  = l_rs1("visdesc")
   end if
 '  response.write(l_existe)
   l_rs1.Close
%>

<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Carga de Vistos de Evaluaci&oacute;n - Evaluaci&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
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

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos">

<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<th colspan="3" align="left" class="th2">Carga de Vistos de Evaluaci&oacute;n</th>
<tr>
<tr style="border-color :CadetBlue;">
	<td>Fecha</td>
	<td>Observaci&oacute;n</td>
	<td>&nbsp;</td>
</tr>
	
<tr>
 <td>
	<b>Firmada el </b>
	<input type="text" name="visfecha" size="10" maxlength="10" value="<%=l_visfecha%>">
	<a href="Javascript:Ayuda_Fecha(document.datos.visfecha)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
</td>
<td>
			<textarea name="visdesc"  maxlength=200 size=200 cols=40 rows=5><%=trim(l_visdesc)%></textarea>
		</td>
	<%if l_existe = "si" then%>	
		<td valign=top><a href=# onclick="if (validarfecha(document.datos.visfecha)) {grabar.location='grabar_vistos_evaluacion_00.asp?tipo=M&evldrnro=<%=l_evldrnro%>&visdesc='+escape(document.datos.visdesc.value)+'&visfecha='+document.datos.visfecha.value;document.datos.grabado.value='M';}">Actualizar</a>			
	<%else%>	
		<td valign=top><a href=# onclick="if (validarfecha(document.datos.visfecha)) {grabar.location='grabar_vistos_evaluacion_00.asp?tipo=A&evldrnro=<%=l_evldrnro%>&visdesc='+escape(document.datos.visdesc.value)+'&visfecha='+document.datos.visfecha.value;document.datos.grabado.value='G';}">Grabar</a>
	<%end if%>	
			<input type="text" readonly disabled name="grabado" size="1">
		</td>
    </tr>
</form>	
</table>
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
