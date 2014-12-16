<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: requerimiento_rel_eyp_00.asp
Descripción: requerimientos solicitado por
Autor : Lisandro Moro
Fecha: 20/04/2004
Modificado  : 12/09/2006 Raul Chinestra - se agregó Requerimientos de Personal en Autogestión   
Modificado:   11/10/2006 Raul Chinestra - Se elimino la posibilidad de ver empleados Inactivos
-->
<% 'on error goto 0
' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma
 
' Filtro
 Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
 Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
 Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)
 
' Orden
 Dim l_Orden      ' Son las etiquetas que aparecen en el orden
 Dim l_CamposOr   ' Son los campos para el orden
 
 Dim rs
 Dim l_rs
 Dim l_sql
 Dim l_ternro
 Dim l_ant_leg
 Dim l_sig_leg
 
 Dim l_empleg
 Dim l_ternom
 Dim l_ternom2
 Dim l_terape
 Dim l_terape2
 Dim blanquear
 
 l_empleg = request("empleg")
 l_ternro = request("ternro")
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 if l_ternro = "" then
	 if l_empleg = "" then
		l_sql = "SELECT empleg, ternro, terape, terape2, ternom, ternom2 "
		l_sql = l_sql & "FROM empleado "
		l_sql = l_sql & "WHERE (empleg = "
		l_sql = l_sql & "(SELECT MIN(empleg) FROM empleado where empleado.empest = -1))"
	 else
		l_sql = "SELECT empleg, ternro, terape, terape2, ternom, ternom2  FROM empleado WHERE empest = -1 and empleg = " & l_empleg
	 end if
 else
 	l_sql = "SELECT empleg, ternro, terape, terape2, ternom, ternom2  FROM empleado WHERE empest = -1 and ternro = " & l_ternro
 end if
 
 l_rs.Maxrecords = 1
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.eof then
 	l_ternro  = l_rs("ternro")
	l_empleg  = l_rs("empleg")
	l_terape  = l_rs("terape")
	l_terape2 = l_rs("terape2")
	l_ternom   = l_rs("ternom")
	l_ternom2  = l_rs("ternom2")	
	blanquear = false
	l_rs.Close	
 else
 	l_ternro  = 0
	blanquear = true
	l_rs.Close
	if l_empleg = "" then
		l_sql = "SELECT empleg, ternro, terape, terape2, ternom, ternom2 "
		l_sql = l_sql & "FROM empleado "
		l_sql = l_sql & "WHERE (empleg = "
		l_sql = l_sql & "(SELECT MIN(empleg) FROM empleado where empleado.empest = -1))"
		 rsOpen l_rs, cn, l_sql, 0
		 if not l_rs.eof then
			l_empleg = l_rs("empleg")
		 else
			l_empleg = 0
		 end if
		 l_rs.Close
	end if
 end if
 
 ' Siguiente/Anterior
 l_sql = "SELECT ternro, empleg FROM empleado WHERE empest = -1 and  empleg < " & l_empleg & " ORDER BY empleg DESC"
' fsql_first l_sql, 1
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.eof then
	l_ant_leg = l_rs("empleg")
 else
	l_ant_leg = l_empleg
 end if
 l_rs.Close
 
 l_sql = "SELECT ternro, empleg FROM empleado WHERE empest = -1 and  empleg > " & l_empleg & " ORDER BY empleg ASC"
' fsql_first l_sql,1
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.eof then
	l_sig_leg = l_rs("empleg")
 else
	l_sig_leg = l_empleg
 end if
 l_rs.Close
 
 if blanquear = true then 
 	l_empleg = ""
	' Abajo se mostrará el cartel correspondiente, una vez cargado el form para que quede mejor visualmente
 end if
 
 
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Solicitado - Empleos y Postulantes - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>
function param(){
	chequear= "cabnro=<%= l_ternro%>";
	return chequear;
}
 	   
function Sig_Ant(leg){
	if (leg != ""){
		document.location ="requerimiento_rel_eyp_00.asp?empleg=" + leg;
	}
}

function nuevoempleado(ternro,empleg,terape,ternom){
	if (empleg != 0){ 
		document.datossol.empleg.value = empleg;
		document.datossol.empleado.value = terape + ", " + ternom;
		Sig_Ant(empleg);
	}
}

function Tecla(num){
	if (num==13) {
		verificacodigo(document.datossol.empleg,document.datossol.empleado,'empleg','terape, ternom','empleado');		
		Sig_Ant(document.datossol.empleg.value);
		return false;
  }
  return num;
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="javascript:document.all.empleg.focus();document.all.empleg.select();">
<form name=datossol>
	<input type="hidden" name="ternro" value="<%= l_ternro %>">

<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr>
	    <td nowrap align="left">
			<a id="antsig" href="JavaScript:Sig_Ant(<%= l_ant_leg %>)">
			<img align="absmiddle" src="/serviciolocal/shared/images/prev.jpg" alt="Empleado Anterior (<%= l_ant_leg %>)" border="0"></a>
			<input type="text" onKeyPress="return Tecla(event.keyCode)" value="<%= l_empleg %>" size="8" name="empleg" onchange="javascript:verificacodigo(this,document.datossol.empleado,'empleg','terape, ternom','empleado');Sig_Ant(this.value);">
			<a id="sigant" href="JavaScript:Sig_Ant(<%= l_sig_leg %>)">
			<img align="absmiddle" src="/serviciolocal/shared/images/next.jpg" alt="Empleado Siguiente (<%= l_sig_leg %>)" border="0"></a>
			<a id="hlp" onclick="JavaScript:window.open('help_emp_01.asp','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no,width=700,height=400');" onmouseover="window.status='Buscar Empleado por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
			<img align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Empleados" border="0">
			</a>
			<input style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_terape & " " & l_terape2 & ", " & l_ternom & " " & l_ternom2%>">
		</td>
	</tr>
</table>
</form>
<%
 if blanquear = true then 
 	' Se coloco esta parte al final del doc para que quede mejor al momento visualmente cuando se muestra el cartel
	Response.write "<script>alert('El Empleado seleccionado no está Activo.');</script>"	
 end if
%>
<script>
	parent.document.datos.reqperrelpor.value = <%= l_ternro%>;
</script>
</body>
</html>

