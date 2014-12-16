<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : registr_diarias_gti_03.asp
Descripcion    : Modulo que se encarga de guardar los datos de una registracion
Modificacion   :
   18/09/2003 - Scarpa D. - Coordinacion con el tablero del empleado
   08/10/2003 - Scarpa D. - Agregado de Proc. OnLine
   08/10/2003 - Scarpa D. - Cerrar la ventana llamadora cuando es una modificacion
   27/10/2003 - Scarpa D. - No Actualizar el tablero cuando se cierra la ventana   
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
on error goto 0

const l_valornulo = "null"

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_regnro
Dim l_regfecha
Dim l_reghora 
Dim l_regentsal
Dim l_regmanual
Dim l_relnro
Dim l_ternro

Dim l_fechadesde
Dim l_fechahasta

Dim l_datos

l_tipo = request.querystring("tipo")
l_fechadesde = request.Form("fechadesde")
l_fechahasta = request.Form("fechahasta")
l_regnro = request.Form("regnro")
l_ternro = request.Form("ternro") 
l_relnro = request.Form("relnro")

if l_relnro = "" then
	l_relnro = l_valornulo
end if
l_regfecha = request.Form("regfecha")
l_regentsal = request.Form("regentsal")
if l_regentsal = "D" then
	l_regentsal = l_valornulo
else 
	l_regentsal = "'" & l_regentsal & "'"
end if
l_reghora = request.Form("reghora1") & request.Form("reghora2")


set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" OR l_tipo = "TA" then 
	l_sql = "insert into gti_registracion "
	l_sql = l_sql & "( ternro, regfecha, reghora, relnro, regentsal, regmanual) "
	l_sql = l_sql & "values (" & l_ternro &", " & cambiafecha(l_regfecha,"YMD",true) & ", '" & l_reghora & "', " & l_relnro 
	l_sql = l_sql & ", " & l_regentsal & ", -1)"
else
	l_sql = "update gti_registracion "
	l_sql = l_sql & "set  ternro="& l_ternro & ", regfecha=" & cambiafecha(l_regfecha,"YMD",true) & ", reghora ='" & l_reghora & "', relnro =" & l_relnro 
	l_sql = l_sql & ", regentsal =" & l_regentsal & ", regmanual= -1"
	l_sql = l_sql & " where regnro = " & l_regnro
end if
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

Set cn = Nothing
Set l_cm = Nothing
l_datos = "ternro="& l_ternro 

Response.write "<script>alert('Operación Realizada.');"
Response.write "window.opener.opener.ifrm.location.reload();"
Response.write "window.opener.close();"
Response.write "window.close();</script>"
	  
%>

</body>
</html>

