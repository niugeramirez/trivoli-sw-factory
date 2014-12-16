<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : ver_objetivos_eva_00.asp
'Objetivo : ver objetivos de evaluacion
'Fecha	  : 27-05-2004
'Autor	  : CCRossi
'Modificacion: 29-12-04-CCRossi- sacar campos para Deloitte
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
 Dim l_rs
 Dim l_sql
 Dim l_filtro
 Dim l_orden

'parametros
 Dim l_evldrnro
 Dim l_evapernro 'periodo de evaluacion
 
 l_evldrnro = request.querystring("evldrnro")
 l_evapernro = request.querystring("evapernro")

 if l_orden = "" then
  l_orden = " ORDER BY evaobjnro "
 end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>

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

function Validar(fecha)
{
	if (fecha == "") {
		alert("Debe ingresar la fecha .");
		return false;
		}
	else
		{
		return true;
		}
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
        <th align=center class="th2">Descripci&oacute;n</th>
        <%if cformed=-1 then%>
        <th align=center class="th2">Forma de Medici&oacute;n</th>
        <%else%>
        <th align=center class="th2">&nbsp;</th>
        <%end if%>
    </tr>
<form name="datos" method="post">
<input type="Hidden" name="terminarsecc" value="SI">
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaobjetivo.evaobjnro,evaperfijo, evapernroeva, evaobjdext,evaobjformed, evldrnro "
l_sql = l_sql & " FROM evaobjetivo "
l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro"
l_sql = l_sql & "		 AND evaluaobj.evaborrador = 0 "
l_sql = l_sql & " WHERE evaluaobj.evldrnro =" & l_evldrnro

rsOpen l_rs, cn, l_sql, 0 

if l_rs.eof then%>
  <tr><td colspan="2"> No se han definido Objetivos.</td></tr>
<%end if

do until l_rs.eof
%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("evaobjnro")%>)">
        <td align=center>
			<textarea readonly disabled name="evaobjdext<%=l_rs("evaobjnro")%>"  cols=70 rows=4><%=trim(l_rs("evaobjdext"))%></textarea>
		</td>
        <td align=center>
			<%if cformed=-1 then%>
			<textarea readonly disabled name="evaobjformed<%=l_rs("evaobjnro")%>"  cols=70 rows=4><%=trim(l_rs("evaobjformed"))%></textarea>
			<%else%>
			<input name="evaobjformed<%=l_rs("evaobjnro")%>" type=hidden value="<%=trim(l_rs("evaobjformed"))%>">
			<%end if%>
			
		</td>
    </tr>
<%	l_rs.MoveNext
loop

l_rs.Close
set l_rs = Nothing

cn.Close
set cn = Nothing
%>

</table>
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

<input type="Hidden" name="cabnro" value="0">
</form>
</body>
</html>
