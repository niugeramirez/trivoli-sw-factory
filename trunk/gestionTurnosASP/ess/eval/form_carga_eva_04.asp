<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<% 
'-------------------------------------------------------------------------------
'Modificado: 20-01-2004-CCRossi-Cambiar select first por la funcion de sqls.inc 
'            14-10-2005 - Leticia Amadio -  Adecuacion a Autogestion	
'		    18-08-2006 - LA. - Sacar la vista v_empleado
'-------------------------------------------------------------------------------
Dim l_sql
Dim l_rs
Dim l_terape 
Dim l_ternom 
Dim l_empleg
Dim l_ternro
Dim l_evaluador
Dim l_evldrnro

Dim l_evacabnro
Dim l_evaseccnro
Dim l_evatevnro

Dim siguiente
Dim Anterior

l_evldrnro = request.querystring("evldrnro")
l_ternro = request.querystring("ternro")

%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Formulario de Carga - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>

<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT evaluador, evacabnro, evatevnro  "
l_sql = l_sql & "FROM evadetevldor  "
l_sql = l_sql & "WHERE evldrnro=" & l_evldrnro
l_rs.Maxrecords = 1
rsOpen l_rs, cn, l_sql, 0
l_evaluador= l_rs("evaluador")
l_evacabnro= l_rs("evacabnro")
l_evatevnro= l_rs("evatevnro")

l_rs.Close
if l_ternro = "" then
	l_ternro= l_evaluador
end if

if l_ternro <> "" then
	l_sql = "SELECT terape, ternom, empleg, ternro "
	l_sql = l_sql & "FROM  empleado "
	l_sql = l_sql & "WHERE ternro=" & l_ternro
	l_rs.Maxrecords = 1
	rsOpen l_rs, cn, l_sql, 0
	l_terape = l_rs("terape")
	l_ternom = l_rs("ternom")
	l_empleg = l_rs("empleg")
else	
	l_empleg= "0"
end if	
l_rs.Close

' Siguiente/Anterior
l_sql = "SELECT empleg FROM empleado where empleg < " & l_empleg & " AND empest = -1 ORDER BY empleg DESC"
l_sql = fsql_first (l_sql,1)
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	anterior = l_rs("empleg")
else
	anterior = l_empleg
end if
l_rs.Close

l_sql = "SELECT empleg,puedever FROM empleado where empleg > " & l_empleg & " AND empest = -1 ORDER BY empleg ASC"
l_sql = fsql_first (l_sql,1)
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	siguiente = l_rs("empleg")
else
	siguiente = l_empleg
end if
l_rs.Close


%>
<script>

function Validar_Formulario(){
var sql="";
if (document.datos.ternro.value==0)	
	alert('Empleado	incorrecto');
else{
	sql= "update evadetevldor set evaluador="+document.datos.ternro.value+" where evacabnro=<%=l_evacabnro%> and evatevnro=<%=l_evatevnro%>";
	abrirVentanaH('form_carga_eva_05.asp?consulta='+sql,'',200,100);
	}
}

function actualizar(){
 opener.actualizar();
 window.close();
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString());
}

function Sig_Ant(leg)
{
if (leg != "")
	{
		abrirVentanaH('nuevo_emp.asp?empleg='+leg,'',200,100);
	}
}


function Tecla(num){
  if (num==13) {
		abrirVentanaH('nuevo_emp.asp?empleg='+document.datos.empleg.value,'',200,100);
		return false;
  }
  return num;
}

function emplerror(nro){
	alert('empleado error:'+nro);
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function nuevoempleado(ternro,empleg,terape,ternom)
{
if (empleg != 0) {	
	document.datos.ternro.value = ternro;
	document.datos.empleg.value = empleg;
	document.datos.empleado.value = terape + ", " + ternom;
	document.location='form_carga_eva_04.asp?evldrnro=<%= l_evldrnro %>&ternro='+ternro;
}
else{
	alert('Empleado	incorrecto');
	document.datos.ternro.value = "0";
	document.datos.empleg.value = "";
	document.datos.empleado.value = "";
	}
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">

<form name="datos" action="" method="post">
<input type="Hidden" name="ternro" value="<%= l_ternro %>">
<%
Const lngAlcanGrupo = 2
dim salir 

cn.Close
set cn = Nothing

%>
<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<th align="left" class="th2">Modificar <%if ccodelco=-1 then%>Rol<%else%>Evaluador<%end if%></th>
    <th align="right" class="th2" valign="middle"> &nbsp;
		<!-- <a class=sidebtnHLP href="Javascript:ayuda('<%'= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a> -->
	</th>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0">
<tr>
	<td colspan="2" height="10"></td>
</tr>
<tr>
    <td nowrap align="right"><b><%if ccodelco=-1 then%>N&uacute;mero<%else%>Empleado<%end if%>:</b></td>
	<td nowrap>
	<a href="JavaScript:Sig_Ant(<%= anterior %>)"><img align="absmiddle" src="/serviciolocal/shared/images/prev.jpg" alt="Anterior (<%= anterior %>)" border="0"></a>
	<input type="text" onKeyPress="return Tecla(event.keyCode)" value="<%= l_empleg %>" size="8" name="empleg">
	<a href="JavaScript:Sig_Ant(<%= siguiente %>)"><img align="absmiddle" src="/serviciolocal/shared/images/next.jpg" alt="Siguiente (<%= siguiente %>)" border="0"></a>
	<a onclick="JavaScript:window.open('help_emp_01.asp','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no,width=700,height=400');" onmouseover="window.status='Buscar Empleado por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
		<img align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda" border="0">
	</a>
	<input style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_terape & ", " &l_ternom%>">

	</td>
</tr>
<tr>
	<td colspan="2" height="10"></td>
</tr>
<tr>
    <td align="right" colspan="2" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>

<% 


%>
</form>	
</body>
</html>
