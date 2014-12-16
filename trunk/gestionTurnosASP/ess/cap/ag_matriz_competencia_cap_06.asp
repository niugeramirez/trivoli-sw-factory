<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: ag_matriz_competencia_cap_06.asp
Descripción: 
Autor : Raul Chinestra
-->
<% 

'Datos del formulario
Dim l_evafacnro
Dim l_origen1
Dim l_origen2

Dim l_competencia
Dim l_fecha
Dim l_porcentaje

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs
Dim l_rs1

%>
<html>
<head>
<link href="../<%= session("estilo")%>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Competencias - Capacitación - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_calendarios.vbs" language="VBScript"></script>

<script>

function Validar_Formulario()
{
if (document.datos.fecha.value == "") {
	alert("Debe ingresar la Fecha.");
	document.datos.fecha.focus();
}	
else
    {
	if (document.datos.evanro.value == 0) {
    	alert("Debe ingresar una Competencia");	
		document.datos.evanro.focus();
	}	
	else
		{
		   if ((document.datos.porcentaje.value > 100)||(document.datos.porcentaje.value < 1)){
		       alert("Debe Ingresar un Porcentaje entre 1 y 100.");document.datos.porcentaje.focus();
	           }else{
		             if (isNaN(document.datos.porcentaje.value))
		                { alert("Debe Ingresar un Porcentaje Numérico.");
		                   document.datos.porcentaje.focus();
		                }
		                else {
		    	               document.datos.submit();
		                     }		
	                 }
         }
	}	 
}

function valido(){
  document.datos.submit();
}

function invalido(texto){
  alert(texto);
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function Ayuda_Fecha(txt)
{ 
  var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

  if (jsFecha == null) txt.value = ''
  else txt.value = jsFecha;

}

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_evafacnro = request.querystring("cabnro")
l_origen1   = request.querystring("origen1")
l_origen2   = request.querystring("origen2")

'response.write l_origen1
'response.write l_origen2

l_sql = "SELECT evafacnro, evafacdesabr, porcen, fecha"
l_sql = l_sql & " FROM evafactor "
l_sql = l_sql & " INNER JOIN cap_capacita ON cap_capacita.entnro = evafactor.evafacnro"
l_sql  = l_sql  & " WHERE evafacnro = " & l_evafacnro & " AND cap_capacita.origen1 =" & l_origen1
l_sql  = l_sql  & " AND cap_capacita.origen2 =" & l_origen2 

rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
   if l_origen1 = "5" and l_origen2 = "3" then
      l_fecha       = l_rs("fecha")
      l_porcentaje  = l_rs("porcen")
	  else Response.write "<script>alert('No se puede Modificar la Competencia ya que no fue cargada de forma Manual.');window.close();</script>"
   end if  
end if

l_rs.Close

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.fecha.focus()">
<form name="datos" action="ag_matriz_competencia_cap_07.asp" method="post" >
<input type="Hidden" name="evafacnro" value="<%= l_evafacnro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2">Datos de la Competencia</td>
	<td colspan="3" class="th2" align="right">		  
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>

<tr>
</tr>

<tr>
    <td align="right"><b>Fecha:</b></td>
	<td colspan="3">
		<input  type="text" name="fecha" size="10" maxlength="10"  value="<%= l_fecha %>"   >
	    <a href="Javascript:Ayuda_Fecha(document.datos.fecha);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
	</td>
</tr>

<tr>
    <td align="right"><b>Competencia:</b></td>
	<td>
		<select style="width:300px" name=evanro size="1">
		  <option value=0>«Seleccione una Opción» </option>
		  
		<%	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")

			l_sql = " SELECT evafacnro, evafacdesabr, evafacdesext "
            l_sql = l_sql & " FROM evafactor "
      	    l_sql = l_sql  & " ORDER BY evafacnro"

            rsOpen l_rs1, cn, l_sql, 0

			do until l_rs1.eof	%>	
			
			<option value= <%= l_rs1("evafacnro") %> > <%= l_rs1("evafacdesabr") %> (<%=l_rs1("evafacnro")%>) </option>
			
			<%	l_rs1.MoveNext
			loop
			l_rs1.Close %>	
		</select>
		<script> document.datos.evanro.value= "<%= l_evafacnro %>"</script>
	</td>	
</tr>
<tr>
    <td align="right"><b>Porcentaje:</b></td>
	<td colspan="3" align="left">
	    <input size="5" type="Text" name="porcentaje" value="<%= l_porcentaje %>" >
	</td>
</tr>

<tr>
    <td  colspan="4" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="blanc.asp" width="100%" height="100%"></iframe> 
</form>
<%
l_rs.close
set l_rs = nothing
'l_Cn.Close
'set l_Cn = nothing
%>
</body>
</html>
