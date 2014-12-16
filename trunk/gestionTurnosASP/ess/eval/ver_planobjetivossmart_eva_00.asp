<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'=====================================================================================
'Archivo  : ver_objetivossmart_eva_00.asp
'Objetivo : ABM de planes para objetivos smart 
'Fecha	  : 18-06-2004
'Autor	  : CCRossi
'Modificar	  : CCRossi - 05-11-2004- crear evaplan de resto de los evaluadores.
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 21-08-2007 - Diego Rosso - Se agrego src="blanc.asp" para https
'=====================================================================================
 Dim l_rs
 Dim l_rs1
 Dim l_cm
 Dim l_sql
 Dim l_filtro
 Dim l_orden

'locales
 dim l_evacabnro 
 dim l_evatevnro 
 dim l_evaluador 
 dim l_planfecharev
 
'parametros
 Dim l_evldrnro
 
 l_evldrnro = request.querystring("evldrnro")

 if l_orden = "" then
  l_orden = " ORDER BY evaobjnro "
 end if

'Crear los evaplan de cada objetivo--------------------------------------------------

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * FROM evadetevldor "
l_sql = l_sql & " WHERE evldrnro = "& l_evldrnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_evacabnro =l_rs("evacabnro")
	l_evatevnro =l_rs("evatevnro")
	l_evaluador =l_rs("evaluador")
end if	
l_rs.Close
Set l_rs = Nothing

'busco el objetivo asociado al mismo evaluador, mismo evatevnro, misma cabecera.
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evadetevldor.evldrnro, evaobjetivo.evaobjnro FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evasecc ON evadetevldor.evaseccnro = evasecc.evaseccnro "
l_sql = l_sql & " INNER JOIN evatiposecc ON evasecc.tipsecnro = evatiposecc.tipsecnro "
l_sql = l_sql & " INNER JOIN evaluaobj   ON evaluaobj.evldrnro=evadetevldor.evldrnro "
l_sql = l_sql & "		 AND evaluaobj.evaborrador = 0 "
l_sql = l_sql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro=evaluaobj.evaobjnro "
l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
'l_sql = l_sql & " AND   evatevnro = " & l_evatevnro
'l_sql = l_sql & " AND   evaluador = " & l_evaluador
l_sql = l_sql & " AND   evadetevldor.evldrnro  <> " & l_evldrnro
l_sql = l_sql & " AND   tipsecobj=-1" 
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0
do while not l_rs.eof 
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT * FROM evaplan "
	l_sql = l_sql & " WHERE evaobjnro = " & l_rs("evaobjnro")
	l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro
'	Response.Write l_sql
	rsOpen l_rs1, cn, l_sql, 0
	if  l_rs1.eof then
		l_rs1.Close
		set l_rs1=nothing
		l_sql= "insert into evaplan (evldrnro,evaobjnro) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_rs("evaobjnro") &")"
'		Response.Write l_sql
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	else
		l_rs1.Close
		set l_rs1=nothing
	end if
	
	l_rs.MoveNext
loop	
l_rs.Close
set l_rs=nothing

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Evaluaci&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
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

 	   
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th align=center class="th2">Objetivos SMART</th>
        <th align=center class="th2">¿Qu&eacute; necesito alcanzar?</th>
        <th align=center class="th2">¿Cu&aacute;les son los pasos <br>que debo tomar para alcanzar mi objetivo?</th>
        <th align=center class="th2">¿Cu&aacute;ndo cumplir&eacute; mi objetivo?</th>
        <th align=center class="th2">¿Qu&eacute; recursos necesito para cumplir mi objetivo?</th>
        <th align=center class="th2">¿Qui&eacute;n brindar&aacute; apoyo para el logro de mi objetivo?</th>
        <th class="th2">&nbsp;</th>
    </tr>
<form name="datos" method="post">
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro, evaobjetivo.evaobjdext,evaplan.aspectomejorar,evaplan.planaccion,"
l_sql = l_sql & " evaplan.planfecharev, evaplan.recursos,evaplan.ayuda, evaplan.evaplnro "
l_sql = l_sql & " FROM evaplan "
l_sql = l_sql & " INNER JOIN evaobjetivo  ON evaobjetivo.evaobjnro = evaplan.evaobjnro"
l_sql = l_sql & " WHERE evaplan.evldrnro =" & l_evldrnro
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
	if trim(l_rs("planfecharev"))="" or isnull(l_rs("planfecharev")) or l_rs("planfecharev")="null" then
		l_planfecharev = date()
	else	
		l_planfecharev = l_rs("planfecharev")
	end if	
	
%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("evaobjnro")%>)">
		<td align=center width=15%>
			<b><%=trim(l_rs("evaobjdext"))%></b>
		</td>
        <td align=center width=20%>
			<textarea readonly disabled name="aspectomejorar<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=20 rows=5><%=trim(l_rs("aspectomejorar"))%></textarea>
		</td>
        <td align=center width=20%>
			<textarea readonly disabled name="planaccion<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=20 rows=5><%=trim(l_rs("planaccion"))%></textarea>
		</td>
		<td nowrap width=10%>
			<input readonly disabled  type="text" name="planfecharev<%=l_rs("evaobjnro")%>" size="10" maxlength="10" value="<%=l_planfecharev%>">
		</td>
        <td align=center width=20%>
			<textarea readonly disabled name="recursos<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=20 rows=5><%=trim(l_rs("recursos"))%></textarea>
		</td>
        <td align=center width=10%>
			<textarea readonly disabled name="ayuda<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=15 rows=5><%=trim(l_rs("ayuda"))%></textarea>
		</td>
        <td valign=top width=5%>
		</td>
    </tr>
<%
	l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>

</table>
<input type="Hidden" name="cabnro" value="0">
<iframe name="grabar"  src="blanc.asp" style="visibility:hidden;width:0;height:0">
<!--iframe name="grabar" style="width:500;height:100"-->


</form>
</body>
</html>
