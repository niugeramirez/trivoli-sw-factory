<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!-----------------------------------------------------------------------------
Archivo     : pol_cuenta_seg_04.asp
Descripcion	: ABM politicas de cuentas de usuarios
Fecha		: 30/07/04
Creador		: Fernando Favre
Modificar	:
-------------------------------------------------------------------------------
-->
<% 
on error goto 0
 Dim l_cm
 Dim l_rs
 Dim l_rs1
 Dim l_sql
 Dim l_pol_nro
 
 l_pol_nro = request.querystring("pol_nro")
 
 Set l_cm = Server.CreateObject("ADODB.Command")
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
 
 l_sql = 		 "SELECT * FROM usr_pol_cuenta "
 l_sql = l_sql & "WHERE pol_nro = " & l_pol_nro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.eof then
	Response.write "<script>alert('Existen usuarios asignados a esta política. No se eliminó el registro.');window.close();</script>"
 else
	l_sql = 		 "SELECT * FROM perf_usr "
	l_sql = l_sql & "WHERE pol_nro = " & l_pol_nro
	rsOpen l_rs1, cn, l_sql, 0
 	if not l_rs1.eof then
		Response.write "<script>alert('Existen perfiles asignados a esta política. No se elimino el registro.');window.close();</script>"
	else
		l_sql = 		"DELETE FROM pol_cuenta " 
		l_sql = l_sql & "WHERE pol_nro = " & l_pol_nro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();// = 'perf_usr_seg_01.asp';window.close();</script>"
	end if
	l_rs1.close
 end if
 l_rs.close
 
 cn.Close
 Set l_cm = Nothing
 Set l_rs = Nothing
 Set l_rs1 = Nothing
 Set cn = Nothing
%>
