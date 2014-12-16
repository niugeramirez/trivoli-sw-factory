<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'----------------------------------------------------------------------------------------------------
' Modificado:13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'----------------------------------------------------------------------------------------------------
%>
<% 
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden

Dim l_evldrnro

l_evldrnro = request.querystring("evldrnro")

if l_orden = "" then
  l_orden = " ORDER BY evaplnro "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Evaluaci&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
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
        <th class="th2">Aspecto a Mejorar</th>
        <th class="th2">Plan de Accion</th>
        <th class="th2">Fecha de Revisión</th>
        <th class="th2">&nbsp;</th>
    </tr>
<form name="datos" method="post">
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaplnro, aspectomejorar, planaccion, planfecharev "
l_sql = l_sql & "FROM evaplan "
l_sql = l_sql & "WHERE evldrnro =" & l_evldrnro
rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("evaplnro")%>)">
        <td>
			<textarea name="aspectomejorar<%=l_rs("evaplnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("aspectomejorar"))%></textarea>
		</td>
        <td>
			<textarea name="planaccion<%=l_rs("evaplnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("planaccion"))%></textarea>
		</td>
        <td>
			<input type="text" name="planfecharev<%=l_rs("evaplnro")%>" size="10" maxlength="10" value="<%=l_rs("planfecharev")%>" readonly>
			<a href="Javascript:Ayuda_Fecha(document.datos.planfecharev<%=l_rs("evaplnro")%>)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
		<td valign=top>
			<a href=# onclick=" if (Validar(document.datos.planfecharev<%=l_rs("evaplnro")%>.value)) {grabar.location='grabar_plan_accion_00.asp?tipo=M&evldrnro=<%=l_evldrnro%>&evaplnro=<%=l_rs("evaplnro")%>&aspectomejorar='+escape(document.datos.aspectomejorar<%=l_rs("evaplnro")%>.value)+'&planaccion='+escape(document.datos.planaccion<%=l_rs("evaplnro")%>.value)+'&planfecharev='+document.datos.planfecharev<%=l_rs("evaplnro")%>.value;document.datos.grabado<%=l_rs("evaplnro")%>.value='M';}">Modificar</a>
			<br>
			<a href=# onclick="grabar.location='grabar_plan_accion_00.asp?tipo=B&evaplnro=<%=l_rs("evaplnro")%>&evldrnro=<%=l_evldrnro%>&aspectomejorar='+document.datos.aspectomejorar<%=l_rs("evaplnro")%>.value+'&planaccion='+document.datos.planaccion<%=l_rs("evaplnro")%>.value+'&planfecharev='+document.datos.planfecharev<%=l_rs("evaplnro")%>.value;document.datos.grabado<%=l_rs("evaplnro")%>.value='B';">Baja</a>
			<br>
			<input type="text" readonly disabled name="grabado<%=l_rs("evaplnro")%>" size="1">
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
    <tr onclick="Javascript:Seleccionar(this,0)">
        <td>
			<textarea name="aspectomejorar"  maxlength=200 size=200 cols=30 rows=4></textarea>
		</td>
        <td>
			<textarea name="planaccion"  maxlength=200 size=200 cols=30 rows=4></textarea>
		</td>
		<td>
			<input type="text" name="planfecharev" size="10" maxlength="10" value="" readonly>
			<a href="Javascript:Ayuda_Fecha(document.datos.planfecharev)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
		<td valign=top><a href=# onclick=" if (Validar(document.datos.planfecharev.value)) {grabar.location='grabar_plan_accion_00.asp?tipo=A&evldrnro=<%=l_evldrnro%>&aspectomejorar='+escape(document.datos.aspectomejorar.value)+'&planaccion='+escape(document.datos.planaccion.value)+'&planfecharev='+document.datos.planfecharev.value;document.datos.grabado.value='G';}">Grabar</a>
		<br>
		<input type="text" readonly disabled name="grabado" size="1">
		</td>
    </tr>

</table>
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

<input type="Hidden" name="cabnro" value="0">
</form>
</body>
</html>
