<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : habilitar_emp_ess_sup_00.asp
Creador        : GdeCos
Fecha Creacion : 4/4/2005
Descripcion    : Pagina encargada de seleccionar los empleados habilitados para el ingreso
				  en Autogestion.
Modificacion   :
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_cm
Dim l_rs
Dim l_sql

Dim l_lista
Dim l_seltipnro

l_seltipnro = 4

'Cargo los datos de los empleados

Set l_rs = Server.CreateObject("ADODB.RecordSet")
	
l_sql = "SELECT empleg,ternro "
l_sql = l_sql & " FROM empleado WHERE empessactivo < 0"
l_sql = l_sql & " ORDER BY empleg "
	
rsOpen l_rs, cn, l_sql, 0 
	
l_lista = "0"
	
do until l_rs.eof 
   l_lista = l_lista & "," & l_rs("ternro") & "@" & l_rs("empleg")
   l_rs.moveNext
loop
	
l_rs.Close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Selecci&oacute;n de Empleados - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>

function guardarEmp(){
//	alert(document.datos.seleccion.value);
  abrirVentanaH('','voculta',100,100);	    
  document.datos.action = 'habilitar_emp_ess_sup_01.asp';
  document.datos.target = 'voculta';
  document.datos.submit();
}

function cerrar(){
  window.close();
}

</script>	
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<form name="datos" action="" method="post">
<input type="Hidden" name="seleccion" value="<%= l_lista%>">
<table  cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
	<td height="100%" width="100%">
		<iframe frameborder="0" scrolling="no" width="100%" height="100%" src="../shared/asp/gen_select_emp_v2_00.asp?seltipnro=<%= l_seltipnro%>&srcdatos=parent.document.datos.seleccion&funcion=parent.guardarEmp&funccerrar=parent.cerrar"></iframe>
	</td>
</tr>
</table>
</form>

</body>
</html>	

