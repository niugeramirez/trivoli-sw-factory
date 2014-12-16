<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<%
'Archivo	: cabaprobada_eva_00.asp
'Descripción: poner cabaprobada en -1
'Autor		: CCRossi
'Fecha		: 01-06-2004
'Modificacion: 
'-------------------------------------------------------------------------------
 Dim l_cm
 Dim l_sql
 
'variables locales
 dim l_hora 
 dim l_arrhr

'parametros de entrada
 dim l_evacabnro 
 dim l_cabaprobada
   
l_evacabnro	= Request.QueryString("evacabnro")
l_cabaprobada	= Request.QueryString("cabaprobada")

function strto2(cad)
	if trim(cad) <>"" then
		if len(cad)<2 then
		'if int(cad)<10 then
			strto2= "0" & cad
		else
			strto2= cad
		end if 
	else
		strto2= "00"
	end if	
end function

' si viene de estadoseccion, viene vacio, le pongo aprobada
if trim(l_cabaprobada)="" or isnull(l_cabaprobada) then
	l_cabaprobada=-1
end if

' ------------------------------------------------------------------------------------
'											BODY 
' ------------------------------------------------------------------------------------
l_hora = mid(time,1,8)
l_arrhr= Split(l_hora,":")
l_hora = strto2(l_arrhr(0))&l_arrhr(1)

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "UPDATE  evacab SET "
l_sql = l_sql & " cabaprobada= "  & l_cabaprobada & ","
l_sql = l_sql & " fechaapro  =   " & cambiafecha(Date(),"","") & ","
l_sql = l_sql & " horaapro   =	'" & l_hora & "'"
l_sql = l_sql & " WHERE evacabnro="& l_evacabnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "UPDATE  evadetevldor SET "
l_sql = l_sql & " habilitado = 0 "
l_sql = l_sql & " WHERE evacabnro="& l_evacabnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

cn.close
Set cn = Nothing

response.write "<script>window.returnValue='0';window.close();</script>"
%>