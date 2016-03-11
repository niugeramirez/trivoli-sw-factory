<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<!--#include virtual="/turnos/shared/inc/adovbs.inc"-->
<!--#include virtual="/turnos/shared/inc/sqls.inc"-->
<% 



on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql

dim l_id
dim l_pacienteid
dim l_apellido
dim l_nombre  
dim l_dni     
dim l_nrohistoriaclinica
dim l_domicilio
dim l_tel
dim l_idobrasocial
dim l_os

dim l_ventana


l_tipo 		     = request.Form("tipo")
l_id             = request.Form("id") 
l_apellido       = request.Form("apellido")
l_nombre         = request.Form("nombre")
l_dni            = request.Form("dni")
l_domicilio      = request.Form("domicilio")
l_nrohistoriaclinica = request.Form("nrohistoriaclinica")
l_tel            = request.Form("tel")
l_idobrasocial   = request.Form("osid")
l_os             = request.Form("os")

l_ventana        = request.Form("ventana") 

if isnull(l_dni) or l_dni = "" then
	l_dni = 0
end if

if isnull(l_nrohistoriaclinica) or l_nrohistoriaclinica = "" then
	l_nrohistoriaclinica = 0
end if

' ------------------------------------------------------------------------------------------------------------------
' codigogenerado() :
' ------------------------------------------------------------------------------------------------------------------
function codigogenerado(tabla)
	Dim l_rs
	Dim l_sql
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = fsql_seqvalue("next_id",tabla)
	rsOpen l_rs, cn, l_sql, 0
	codigogenerado=l_rs("next_id")
	l_rs.Close
	Set l_rs = Nothing
end function 'codigogenerado()


'Al operar sobre varias tablas debo iniciar una transacción
 cn.BeginTrans

set l_cm = Server.CreateObject("ADODB.Command")

if l_tipo = "A" then
	if l_dni <> "" then
		l_sql = "INSERT INTO clientespacientes "
		l_sql = l_sql & " (apellido, nombre, dni,domicilio, telefono, idobrasocial, nrohistoriaclinica, empnro,created_by,creation_date,last_updated_by,last_update_date)"
		l_sql = l_sql & " VALUES ('" & l_apellido & "','" & l_nombre & "'," & l_dni & ",'" & l_domicilio & "','" & l_tel & "'," & l_idobrasocial & ",'" & l_nrohistoriaclinica & "','" & session("empnro") & "','" & session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
	else
		l_sql = "INSERT INTO clientespacientes "
		l_sql = l_sql & " (apellido, nombre, domicilio, telefono, idobrasocial, nrohistoriaclinica, empnro,created_by,creation_date,last_updated_by,last_update_date)"
		l_sql = l_sql & " VALUES ('" & l_apellido & "','" & l_nombre & "','" & l_domicilio & "','" & l_tel & "'," & l_idobrasocial & ",'" & l_nrohistoriaclinica & "','" & session("empnro") & "','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
	
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0	
	
	'Ingreso la lista de empleados a la tabla
	l_id = codigogenerado("clientespacientes")	
	
	'l_sql = " SELECT @@IDENTITY AS 'Identity' "
	'l_cm.activeconnection = Cn
	'l_cm.CommandText = l_sql
	'cmExecute l_cm, l_sql, 0		
	
	Set l_cm = Nothing	
	
	
else

	l_sql = "UPDATE clientespacientes "
	l_sql = l_sql & " SET apellido    = '" & l_apellido & "'"
	l_sql = l_sql & "    ,nombre    = '" & l_nombre & "'"
	l_sql = l_sql & "    ,nrohistoriaclinica    = '" & l_nrohistoriaclinica & "'"	
	if l_dni <> "" then
		l_sql = l_sql & "    ,dni    =    " & l_dni & ""
	end if
	l_sql = l_sql & "    ,domicilio     = '" & l_domicilio & "'"
	l_sql = l_sql & "    ,telefono      = '" & l_tel & "'"	
	l_sql = l_sql & "    ,idobrasocial      = " & l_idobrasocial
	l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
	l_sql = l_sql & "    ,last_update_date = GETDATE()" 	
	l_sql = l_sql & " WHERE id = " & l_id
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0	
	Set l_cm = Nothing
	
end if	
	



cn.CommitTrans 


'Response.write "OK"
Response.write "[{""resultado"":""OK"",""id"":""" & l_id & """}]"

%>

