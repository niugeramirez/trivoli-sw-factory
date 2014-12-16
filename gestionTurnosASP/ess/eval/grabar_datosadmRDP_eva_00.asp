<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'--------------------------------------------------------------------------
'Archivo       : grabar_datosadmRDP_eva_00.asp 
'Descripcion   : Alta/ Modific de los datos de evaadmdatos
'Creacion      : 14-02-2005 
'Autor         : Leticia Amadio.
'Modificacion  : 
'--------------------------------------------------------------------------


on error goto 0 
' variables 
' parametros de entrada ----------------------------------------
  Dim l_evldrnro 
  dim l_tipo 
  
  Dim l_horas 
  Dim l_fechareunion 
  Dim l_basereunion  

'locales 
 dim	l_evacabnro 
' variables de base de datos ------------------------------------
  Dim l_cm 
  Dim l_sql
  Dim l_rs 
  
' parametros de entrada
  l_evldrnro		= request.querystring("evldrnro")
  l_horas			= request("horas")
  l_fechareunion	= request("fechareunion")
  l_basereunion		= request("basereunion")
  
  l_tipo			= request.querystring("Tipo")

'Response.Write l_horas & "<br>"
'Response.Write l_fechareunion & "<br>"
'Response.Write l_basereunion & "<br>"

'BODY ----------------------------------------------------------
'Buscar el nro de cabecera de esta evaluacion
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evacabnro "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " WHERE evadetevldor.evldrnro = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
   l_evacabnro= l_rs("evacabnro")
 end if 
 l_rs.close 
 set l_rs=nothing

if l_tipo="M" then

	'Buscar el nro de cabecera de esta evaluacion
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evadetevldor.evldrnro "
	l_sql = l_sql & " FROM  evadetevldor "
	l_sql = l_sql & " INNER JOIN evadatosadm ON evadetevldor.evldrnro=evadatosadm.evldrnro" 
	l_sql = l_sql & " WHERE evadetevldor.evacabnro = " & l_evacabnro
	rsOpen l_rs, cn, l_sql, 0
	do while not l_rs.EOF 
		l_sql = "UPDATE evadatosadm SET "
		l_sql = l_sql & " horas = NULL,"
		l_sql = l_sql & " fechareunion = " & cambiafecha(l_fechareunion,"YMD",false) & ","
		l_sql = l_sql & " basereunion = " & l_basereunion
		l_sql = l_sql & " WHERE evadatosadm.evldrnro="  & l_rs("evldrnro")
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		l_rs.MoveNext
	loop 
	l_rs.close 
	set l_rs=nothing 
	
else ' es ALTA 
	l_sql = "INSERT INTO evadatosadm "
	l_sql = l_sql & "(evldrnro,horas,fechareunion,basereunion)"
	l_sql = l_sql & " VALUES (" & l_evldrnro & "," & l_horas 
	l_sql = l_sql &  		 "," & cambiafecha(l_fechareunion,"YMD",false) & "," & l_basereunion & ")"
	'response.write "<script> alert(" & l_sql & ") </script>"
	
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end if 
%>
