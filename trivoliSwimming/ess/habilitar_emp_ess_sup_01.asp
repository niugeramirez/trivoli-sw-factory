<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : habilitar_emp_ess_sup_01.asp
Creador        : GdeCos
Fecha Creacion : 4/4/2005
Descripcion    : Pagina encargada de seleccionar los empleados habilitados para el ingreso
				  en Autogestion.
Modificacion   :
-----------------------------------------------------------------------------
-->
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<% 
on error goto 0

Dim l_cm
Dim l_rs
Dim l_rs2
Dim l_sql

Dim l_seleccion

l_seleccion = request.form("seleccion")


cn.beginTrans

'Guardo los datos en la BD

Set l_rs = Server.CreateObject("ADODB.RecordSet")	
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")	
Set l_cm = Server.CreateObject("ADODB.Command")

l_cm.activeconnection = Cn	

Dim l_arr
Dim l_arr2
Dim l_i
Dim l_ternro

l_arr = Split(l_seleccion,",")		

l_seleccion = "0"

for l_i = 1 to UBound(l_arr)
   l_arr2 = split(l_arr(l_i),"@")
   l_seleccion = l_seleccion & "," & l_arr2(0)
next

l_arr = Split(l_seleccion,",")		
l_i = 1

do while l_i <= UBound(l_arr)
	l_ternro = l_arr(l_i)

    'Inserto el dato en la BD
    l_sql = "UPDATE empleado SET empessactivo = -1 "
    l_sql = l_sql & " WHERE ternro = " & l_ternro 
	
    l_cm.CommandText = l_sql
    cmExecute l_cm, l_sql, 0
		  	
	l_i = l_i + 1
loop
	
'borro todos los empleados que no quedaron en la lista
l_sql = "UPDATE empleado SET empessactivo = 0 "
l_sql = l_sql & " WHERE empessactivo = -1 "
if l_seleccion <> "0" then
   l_sql = l_sql & " AND ternro NOT IN ("  & l_seleccion & ") "
end if

l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

cn.commitTrans

%>
<script>
	   alert('Operación Realizada.');
	   window.close();
	   opener.window.close();
</script>


