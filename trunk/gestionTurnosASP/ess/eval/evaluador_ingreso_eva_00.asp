<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<%
'Archivo	: evaluador_ingreso_eva_00.asp
'Descripción: poner evadetevldor.ingreso=-1
'Autor		: CCRossi
'Fecha		: 10-06-2004
'Modificacion: 
'-------------------------------------------------------------------------------
 Dim l_cm
 Dim l_sql
 Dim l_rs
 
'variables locales
 dim l_hora 
 dim l_arrhr
 dim l_ingreso
 
'parametros de entrada
 dim l_evldrnro 
  
l_evldrnro	= Request.QueryString("evldrnro")
 
function strto2(cad)
	if trim(cad) <>"" then
		if len(cad)<2 then
			strto2= "0" & cad
		else
			strto2= cad
		end if 
	else
		strto2= "00"
	end if	
end function

' ------------------------------------------------------------------------------------
'											BODY 
' ------------------------------------------------------------------------------------
l_hora = mid(time,1,8)
l_arrhr= Split(l_hora,":")
l_hora = strto2(l_arrhr(0))& l_arrhr(1)
	
l_ingreso=0

Set l_rs = Server.CreateObject("ADODB.RecordSet")		
l_sql = "SELECT ingreso "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " WHERE evldrnro = " & l_evldrnro
rsOpen l_rs, cn, l_sql, 0 
if not  l_rs.eof then
	l_ingreso = l_rs("ingreso")
end if
l_rs.close	
set l_rs=nothing

if l_ingreso<>-1 then
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "UPDATE  evadetevldor SET "
	l_sql = l_sql & " ingreso =-1,"  
	l_sql = l_sql & " fechaing  =   " & cambiafecha(Date(),"","") & ","
	l_sql = l_sql & " horaing   =	'" & l_hora & "'"
	l_sql = l_sql & " WHERE evldrnro="& l_evldrnro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end if
			
cn.close
Set cn = Nothing

response.write "<script>window.close();</script>"
%>