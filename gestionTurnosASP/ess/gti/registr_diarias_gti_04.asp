<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : registr_diarias_gti_04.asp
Descripcion    : Modulo que se encarga de eliminar los datos de una registracion
Modificacion   :
   08/10/2003 - Scarpa D. - Agregado de Proc. OnLine
-----------------------------------------------------------------------------
-->
<html>
<head>
</head>
<body>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>

<% 
'on error goto 0

Dim l_cm
Dim l_rs
Dim l_sql
Dim l_emp_licnro
Dim l_filtro
Dim l_ternro
Dim l_regnro
Dim l_regmanual

Dim l_fechadesde
Dim l_fechahasta
Dim l_datos

l_regnro = request.querystring("cabnro")
l_ternro = request.querystring("ternro")
l_fechadesde = request.querystring("fechadesde")
l_fechahasta = request.querystring("fechahasta")

'Busco cual es el rango de fechas de la registracion
Dim l_desde
Dim l_hasta
Dim l_empternro

l_sql = " SELECT * FROM gti_registracion " 
l_sql = l_sql & "WHERE regnro = " & l_regnro

Set l_rs = Server.CreateObject("ADODB.RecordSet")

rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
  l_desde     = l_rs("regfecha")
  l_hasta     = l_rs("regfecha")
  l_empternro = l_rs("ternro")
end if

l_rs.close


set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "DELETE FROM gti_registracion " 
	l_sql = l_sql & "WHERE regnro = " & l_regnro
	
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

l_datos = "ternro="& l_ternro & "&fechadesde="& l_fechadesde & "&fechahasta="& l_fechahasta

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>

</body>
</html>

