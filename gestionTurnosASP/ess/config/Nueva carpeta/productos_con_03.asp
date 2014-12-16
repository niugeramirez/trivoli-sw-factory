<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<%
'Archivo: productos_con_03.asp
'Descripci�n: Habilitaci�n/Deshabilitaci�n de Productos
'Autor : Javier Posadas
'Fecha: 05/04/2005

Dim l_tipo
Dim l_rs
Dim l_cm
Dim l_sql

Dim l_pronro
Dim l_proest
Dim l_profull
Dim l_a

l_profull = request.QueryString("cabnro")
l_tipo	  = request.QueryString("tipo")
l_pronro  = split(l_profull,",")

if ( trim(cstr(l_tipo)) = "H" or trim(cstr(l_tipo)) = "HT" ) then
	l_proest = -1
else
	l_proest = 0
end if

set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql

'Habilitaci�n/Deshabilitaci�n de todos los Productos

if ( trim(cstr(l_tipo)) = "DT" or trim(cstr(l_tipo)) = "HT" ) then
	
	'Inicio la Transacci�n
	Cn.BeginTrans
	
	l_sql = " UPDATE tkt_producto "
	l_sql = l_sql & " SET proest = " & l_proest 
	
	cmExecute l_cm, l_sql, 0
	
	'Cierro la Transacci�n
	Cn.CommitTrans
else
	'Habilitaci�n/Deshabilitaci�n de los Productos seleccionados
	
	'Inicio la Transacci�n
	Cn.BeginTrans

	for l_a = 0 to UBound(l_pronro)
		l_sql = " UPDATE tkt_producto "
		l_sql = l_sql & " SET proest = " & l_proest 
		l_sql = l_sql & " WHERE pronro = " & l_pronro(l_a)
		
		cmExecute l_cm, l_sql, 0
	next
	
	'Cierro la Transacci�n
	Cn.CommitTrans
end if

Set l_cm = Nothing

Response.write "<script>alert('Operaci�n Realizada.');window.parent.ifrm.location.reload();</script>"
%>
