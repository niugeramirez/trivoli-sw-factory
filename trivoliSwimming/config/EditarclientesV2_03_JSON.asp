<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/adovbs.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/sqls.inc"-->
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

dim l_telefono
dim l_celular  
dim l_mail  
dim l_direccion
dim l_idciudad


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


l_telefono		           = request.Form("telefono")
l_celular		           = request.Form("celular")
l_mail  		           = request.Form("mail")
l_direccion		           = request.Form("direccion")
l_idciudad				   = request.Form("idciudad")

l_ventana        = request.Form("ventana") 

if isnull(l_dni) or l_dni = "" then
	l_dni = 0
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

'response.write "l_tipo "&l_tipo

if l_tipo = "A" then

	l_sql = "INSERT INTO clientes "
	l_sql = l_sql & " (nombre, telefono, celular,mail, direccion, idciudad, empnro,created_by,creation_date,last_updated_by,last_update_date)"
	l_sql = l_sql & " VALUES ('" & l_nombre & "','" &  l_tel & "','" & l_celular & "','" & l_mail & "','" & l_direccion & "'," & l_idciudad & ",'" & session("empnro") & "','" & session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"

	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0	
	
	'Ingreso la lista de empleados a la tabla
	l_id = codigogenerado("clientes")	
	
	'l_sql = " SELECT @@IDENTITY AS 'Identity' "
	'l_cm.activeconnection = Cn
	'l_cm.CommandText = l_sql
	'cmExecute l_cm, l_sql, 0		
	
	Set l_cm = Nothing		
else

	l_sql = "UPDATE clientes "
	l_sql = l_sql & " SET nombre    = '" & l_nombre & "'"
		l_sql = l_sql & "    ,telefono  = '" & l_telefono & "'"	
		l_sql = l_sql & "    ,celular   = '" & l_celular & "'"
		l_sql = l_sql & "    ,mail      = '" & l_mail & "'"
		l_sql = l_sql & "    ,direccion = '" & l_direccion & "'"
		l_sql = l_sql & "    ,idciudad  =  " & l_idciudad & ""			
	
	'if l_dni <> "" then
	'	l_sql = l_sql & "    ,dni    =    " & l_dni & ""
	'end if
	'l_sql = l_sql & "    ,domicilio     = '" & l_domicilio & "'"
	'l_sql = l_sql & "    ,telefono      = '" & l_tel & "'"	
	'l_sql = l_sql & "    ,idobrasocial      = " & l_idobrasocial
	l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
	l_sql = l_sql & "    ,last_update_date = GETDATE()" 	
	l_sql = l_sql & " WHERE id = " & l_id
	' response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0	
	Set l_cm = Nothing
	
end if	
	



cn.CommitTrans 

'response.write "l_id "&l_id
'Response.write "OK"
Response.write "[{""resultado"":""OK"",""id"":""" & l_id & """}]"

%>

