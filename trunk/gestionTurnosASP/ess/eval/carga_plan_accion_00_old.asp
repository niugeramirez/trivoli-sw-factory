<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion

' Variables
' de parametros entrada
  
' de uso local  
  
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evldrnro = Request.QueryString("evldrnro")
  
 'HARCODED----------------------
  'l_evldrnro = 2
   
' Crear registros de evaNOTAS para evldrnro y el tipo nota
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT evldrnro, evatevnro  "
  l_sql = l_sql & "FROM  evadetevldor "
  l_sql = l_sql & "ORDER BY evatevnro "
  rsOpen l_rs, cn, l_sql, 0

  do while not l_rs.EOF 
	    Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	    l_sql = "SELECT *  "
        l_sql = l_sql & "FROM  evaplan "
        l_sql = l_sql & "WHERE evaplan.evldrnro   = " & l_rs("evldrnro")
        rsOpen l_rs1, cn, l_sql, 0

        if l_rs1.EOF then
			l_sql = "INSERT INTO evaplan "
	        l_sql = l_sql & "(evldrnro) "
			l_sql = l_sql & " VALUES (" & l_rs("evldrnro") & ")"
		end if

		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		l_rs1.Close
		
		l_rs.MoveNext
 loop
 l_rs.close	
%>

<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Planes de Acci&oacute;n - Evaluaci&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

//function Recargar01()
//{
//	document.ifrm.location.href= 'competencias_01.asp?evatitnro=' + document.datos.evatitnro.value;
//}
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

function Validar(fecha)
{
	if (fecha == "") 
		alert("Debe ingresar la fecha .");
	else
		{
		return false;
		}
}
 	   
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos">

<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<td colspan="4" align="left" class="th2">Planes de Acci&oacute;n</td>
<tr>
<tr style="border-color :CadetBlue;">
	<td>Aspecto a Mejorar</td>
	<td>Plan de Accion</td>
	<td>Fecha</td>
	<td>&nbsp;</td>
</tr>
	
<%'BUSCAR evavistos
   Set l_rs = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evldrnro, aspectomejorar, planaccion, planfecharev "
   l_sql = l_sql & "FROM evaplan "
   l_sql = l_sql & "WHERE evaplan.evldrnro      = " & l_evldrnro
   l_sql = l_sql & "ORDER BY evaplan.planfecharev " 
   rsOpen l_rs, cn, l_sql, 0
   do while not l_rs.eof %>
   <tr>
		<td>
		<textarea name="aspectomejorar<%=l_rs("evldrnro")%>"  maxlength=200 size=200 cols=40 rows=5>
			<%=trim(l_rs("aspectomejorar"))%>	
		</textarea>
		</td>
		
		<td>
		<textarea name="planaccion<%=l_rs("evldrnro")%>"  maxlength=200 size=200 cols=40 rows=5><%=trim(l_rs("planaccion"))%></textarea>
		</td>
		
		<td valign=top>
			<input type="text" name="planfecharev<%=l_rs("evldrnro")%>" size="10" maxlength="10" value="<%=l_rs("planfecharev")%>" readonly>
			<a href="Javascript:Ayuda_Fecha(document.datos.planfecharev<%=l_rs("evldrnro")%>)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
		<td valign=top><a href=# onclick="Validar(document.datos.planfecharev<%=l_rs("evldrnro")%>.value);grabar.location='grabar_plan_accion_00.asp?evldrnro=<%=l_evldrnro%>&aspectomejorar='+document.datos.aspectomejorar<%=l_rs("evldrnro")%>.value+'&planaccion='+document.datos.planaccion<%=l_rs("evldrnro")%>.value+'&planfecharev='+document.datos.planfecharev<%=l_rs("evldrnro")%>.value;document.datos.grabado<%=l_rs("evldrnro")%>.value='G';">Grabar</a>
		<input type="text" readonly disabled name="grabado<%=l_rs("evldrnro")%>" size="1">
		</td>
    </tr>
  <%l_rs.Movenext
  loop
  l_rs.Close%>
	

<iframe name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>
</form>	
</table>

</body>
</html>
