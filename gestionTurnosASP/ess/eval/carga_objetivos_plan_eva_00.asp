<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'___________________________________________________________________________________
'Archivo   : carga_objetivos_plan_eva_00.asp
'Objetivo  : carga de objetivos con Plan de Desarrollo 
'Fecha	   : 12-12-2006
'Autor	   : Leticia Amadio
'Modificado: 15-12-2006 - LA -  Adecuacion a Autogestion
'				24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'___________________________________________________________________________________

on error goto 0

 Dim l_rs
 Dim l_sql
 Dim l_filtro
 Dim l_orden

'parametros
Dim l_evldrnro
Dim l_evapernro 'periodo de evaluacion
Dim l_cantidad ' variable que contiene el numero maximo de caracteres permitido en el Texto
Dim l_cantidad2 'el numero max d ecaracteres permitidos para la descipciones de plan de desarrollo y resultados esperados

l_evldrnro = request.querystring("evldrnro")
l_evapernro = request.querystring("evapernro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
if l_orden = "" then
  l_orden = " ORDER BY evaobjnro "
end if

l_cantidad= 250   
l_cantidad2= 1500 
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/texto_texttarea.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>

function Ayuda_Fecha(txt){
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto){
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

  

var jsSelRow = null;

function Deseleccionar(fila){
 fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro){
 if (jsSelRow != null) {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}


function GrabarObjetivo(tipo,evldrnro,evapernro,evaobjnro,evaobjdext,letra){
var aux

if (eval('document.datos.evaobjdext'+evaobjnro+'.value')==""){
	alert('La descripción no puede ser vacia.');
	eval('document.datos.evaobjdext'+evaobjnro+'.focus()');
	return
}

if (!stringValido(eval('document.datos.evaobjdext'+evaobjnro+'.value'))){
	alert('La Descripción contiene caracteres no válidos.');
	eval('document.datos.evaobjdext'+evaobjnro+'.focus()');
	return;
}

aux=eval('document.datos.evaobjplan'+evaobjnro+'.value'); 
document.datos.evaobjplan.value = aux; 

if (!stringValido(aux)){
	alert('El Plan de Desarrollo contiene caracteres no válidos.');
	eval('document.datos.evaobjplan'+evaobjnro+'.focus()');
	return;
}
  
aux=eval('document.datos.evaobjresu'+evaobjnro+'.value');
document.datos.evaobjresu.value = aux; 
if (!stringValido(aux)){
	alert('Los Resultados esperados contiene caracteres no válidos.');
	eval('document.datos.evaobjresu'+evaobjnro+'.focus()');
	return;
}

aux=eval('document.datos.evaobjfecha'+evaobjnro+'.value');
document.datos.evaobjfecha.value = aux; 
	
if ( aux !== "" && (!validarfecha(eval('document.datos.evaobjfecha'+evaobjnro))) ){ 

} else {
	aux=eval('document.datos.grabado' +evaobjnro+'.value="'+letra+'"');
	document.datos.target ="grabar";
	document.datos.method ="post";
	document.datos.action = 'grabar_objetivos_eva_00.asp?tipo='+ tipo 
							+'&evldrnro='+evldrnro
							+'&evapernro='+evapernro
							+'&evaobjnro='+evaobjnro
							+'&evaobjdext='+evaobjdext;
	document.datos.submit();
}

}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
<tr>
	<th colspan=2 class="th2">Descripci&oacute;n</th>
	<th class="th2">&nbsp;</th>
</tr>
<form name="datos" method="post">
<input type="Hidden" name="terminarsecc" value="SI">

<%
l_sql = "SELECT evaobjetivo.evaobjnro, evaobjdext,evldrnro, evaobjplan, evaobjresu, evaobjfecha  "
l_sql = l_sql & " FROM evaobjetivo "
l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro "
l_sql = l_sql & " LEFT JOIN evaobjplan ON evaobjplan.evaobjnro=evaobjetivo.evaobjnro "
l_sql = l_sql & " WHERE evaluaobj.evldrnro =" & l_evldrnro
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0 

do until l_rs.eof 
%>
<tr onclick="Javascript:Seleccionar(this,<%= l_rs("evaobjnro")%>)">
	<td rowspan=4 valign=top align=right>
		<br><strong>OBJETIVO: </strong> &nbsp;&nbsp;<br>
	</td>	
    <td nowrap> 
		&nbsp;<br>
		<textarea name="evaobjdext<%=l_rs("evaobjnro")%>"  size=200 cols=75 rows=3 onKeyUp="Limite(this.value,<%=l_cantidad%>);Contador(document.datos.contador<%=l_rs("evaobjnro")%>,this,<%=l_cantidad%>);"><%=trim(l_rs("evaobjdext"))%></textarea>
		<input CLASS="rev" disabled style="background:#e0e0de;" readonly type=text name="contador<%=l_rs("evaobjnro")%>" size=2>	
		<script> document.datos.contador<%=l_rs("evaobjnro")%>.value = <%=l_cantidad%> - document.datos.evaobjdext<%=l_rs("evaobjnro")%>.value.length </script>
		<br>&nbsp;
	</td>
    <td valign=top align=center rowspan=4>
		&nbsp;<br>&nbsp;<br>
		<a href=# onclick="javascript:GrabarObjetivo('MPL',<%=l_evldrnro%>,<%=l_evapernro%>,<%=l_rs("evaobjnro")%>,document.datos.evaobjdext<%=l_rs("evaobjnro")%>.value,'M');">Modificar</a>
		&nbsp;&nbsp;&nbsp;<br>&nbsp;<br>
		<a href=# onclick="javascript:GrabarObjetivo('BPL',<%=l_evldrnro%>,<%=l_evapernro%>,<%=l_rs("evaobjnro")%>,document.datos.evaobjdext<%=l_rs("evaobjnro")%>.value,'B');">Eliminar Obj.</a>
		<br>
		<input type="text" readonly disabled name="grabado<%=l_rs("evaobjnro")%>" size="1" value="G">
	</td>
</tr>
<tr>
	<td>
		<strong>Plan de Desarrollo: </strong><br>
		<textarea name="evaobjplan<%=l_rs("evaobjnro")%>"  cols=85 rows=4 onKeyUp="Limite(this.value,<%=l_cantidad2%>);Contador(document.datos.contadorp<%=l_rs("evaobjnro")%>,this,<%=l_cantidad2%>);"><%=trim(l_rs("evaobjplan"))%></textarea>
		<input CLASS="rev" disabled style="background:#e0e0de;" readonly type=text name="contadorp<%=l_rs("evaobjnro")%>" size=2>	
		<script> document.datos.contadorp<%=l_rs("evaobjnro")%>.value = <%=l_cantidad2%> - document.datos.evaobjplan<%=l_rs("evaobjnro")%>.value.length </script>			
	</td>	
</tr>
<tr>
	<td>
		<strong>Resultados esperados:</strong> <br>
		<textarea name="evaobjresu<%=l_rs("evaobjnro")%>"  cols=85 rows=4 onKeyUp="Limite(this.value,<%=l_cantidad2%>);Contador(document.datos.contadorr<%=l_rs("evaobjnro")%>,this,<%=l_cantidad2%>);"><%=trim(l_rs("evaobjresu"))%></textarea>
		<input CLASS="rev" disabled style="background:#e0e0de;" readonly type=text name="contadorr<%=l_rs("evaobjnro")%>" size=2>
		<script> document.datos.contadorr<%=l_rs("evaobjnro")%>.value = <%=l_cantidad2%> - document.datos.evaobjresu<%=l_rs("evaobjnro")%>.value.length </script>			
	</td>
</tr>
<tr>
	<td>
		<strong>Fecha:</strong> <br>
		<input type="text" name="evaobjfecha<%=l_rs("evaobjnro")%>" value="<%=l_rs("evaobjfecha")%>" size="10" maxlength="10">&nbsp;
		<a href="Javascript:Ayuda_Fecha(document.datos.evaobjfecha<%=l_rs("evaobjnro")%>)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
 		<br>&nbsp;
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

</tr>
	<td colspan=3>&nbsp;&nbsp;</td>
<tr>
<tr onclick="Javascript:Seleccionar(this,0)">
	<td rowspan=4 valign=top align=right>
		<br>
		<strong>OBJETIVO&nbsp;&nbsp;<br>&nbsp;&nbsp;&nbsp;&nbsp;NUEVO: </strong> &nbsp;&nbsp;<br>
	</td>	
	<td nowrap> 
	&nbsp;<br>
	<textarea name="evaobjdext0" size=200 cols=75 rows=3 onKeyUp="Limite(this.value,<%=l_cantidad%>);Contador(document.datos.contador0,this,<%=l_cantidad%>);"></textarea>
	<input CLASS="rev" disabled style="background:#e0e0de;" readonly type=text name="contador0" size=2>
	<script> document.datos.contador0.value = <%=l_cantidad%> - document.datos.evaobjdext0.value.length </script>
	<br>&nbsp;
	</td>
	<td valign=top rowspan=4 align=center>
		&nbsp;<br>&nbsp;<br>
		<a href=# onclick="javascript:GrabarObjetivo('PL',<%=l_evldrnro%>,<%=l_evapernro%>,0,document.datos.evaobjdext0.value,'G');">Grabar</a>
		<!--document.datos.grabado.value='G'; }--> 
		<br>
		<input type="text" readonly disabled name="grabado0" size="1">
</td>
</tr>
<tr>
	<td> 
		<strong>Plan de Desarrollo: </strong><br>
		<textarea name="evaobjplan0"  cols=85 rows=4 onKeyUp="Limite(this.value,<%=l_cantidad%>);Contador(document.datos.contadorp0,this,<%=l_cantidad%>);"></textarea>
		<input CLASS="rev" disabled style="background:#e0e0de;" readonly type=text name="contadorp0" size=2>			
		<script> document.datos.contadorp0.value = <%=l_cantidad2%> - document.datos.evaobjplan0.value.length </script>			
	</td>
</tr>
<tr>
	<td>
	 <strong>Resultados esperados:</strong> <br>
		<textarea name="evaobjresu0"  cols=85 rows=4 onKeyUp="Limite(this.value,<%=l_cantidad%>);Contador(document.datos.contadorr0,this,<%=l_cantidad%>);"></textarea>
		<input CLASS="rev" disabled style="background:#e0e0de;" readonly type=text name="contadorr0" size=2>
		<script> document.datos.contadorr0.value = <%=l_cantidad2%> - document.datos.evaobjresu0.value.length </script>			
	</td>
</tr>
<tr>
	<td>
		<strong>Fecha:</strong> <br>
		<input type="text" name="evaobjfecha0" value="" size="10"  maxlength="10">&nbsp;
		<a href="Javascript:Ayuda_Fecha(document.datos.evaobjfecha0)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
	</td>	
</tr>
</table>
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
<!--iframe name="grabar" style="width:500;height:100" -->
</iframe>
<input type="Hidden" name="evaobjplan" value="">
<input type="Hidden" name="evaobjresu" value="">
<input type="Hidden" name="evaobjfecha" value="">
<input type="Hidden" name="cabnro" value="0">
</form>
</body>
</html>
