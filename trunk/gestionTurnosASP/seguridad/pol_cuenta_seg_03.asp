<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
 Dim l_cm
 Dim l_rs 'para realizar una consulta de control
 Dim l_sql
 Dim l_pol_nro
 Dim l_pass_expira_dias
 Dim l_pass_camb_dias
 Dim l_pass_longitud
 Dim l_pass_historia
 Dim l_pol_desc 
 Dim l_pass_int_fallidos
 Dim l_pass_dias_log
 Dim l_pass_cambiar
 Dim l_pass_log_simult
 
 Dim l_tipo
 
 l_tipo				 = request.QueryString("tipo")
 l_pol_nro			 = request.Form("pol_nro")
 l_pass_expira_dias  = request.Form("pass_expira_dias")
 l_pass_camb_dias	 = request.Form("pass_camb_dias")
 l_pass_longitud	 = request.Form("pass_longitud")
 l_pass_historia	 = request.Form("pass_historia")
 l_pol_desc			 = request.Form("pol_desc")
 l_pass_int_fallidos = request.Form("pass_int_fallidos")
 l_pass_dias_log	 = request.Form("pass_dias_log")
 l_pass_cambiar		 = request.Form("pass_cambiar")
 l_pass_log_simult	 = request.Form("pass_log_simult")
 
 if len(l_pass_cambiar) > 0 then l_pass_cambiar = -1 	else l_pass_cambiar = 0 	end if
 
 l_pass_log_simult = 0
 
 set l_cm = Server.CreateObject("ADODB.Command")
 set l_rs = Server.CreateObject("ADODB.RecordSet")
 
 l_sql = 		 "SELECT pol_nro "
 l_sql = l_sql & "FROM pol_cuenta "
 l_sql = l_sql & "WHERE pol_desc = '" & l_pol_desc & "'"
 if l_tipo = "M" then
 	l_sql = l_sql & " AND pol_nro <> " & l_pol_nro
 end if
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.eof then
	Response.write "<script>window.history.back();alert('Ya existe una Política con esta Descripción.');</script>"
 else
	 if l_tipo = "A" then
		l_sql =			"INSERT INTO pol_cuenta (pol_desc, pass_expira_dias, pass_camb_dias, pass_longitud, pass_historia, "
		l_sql = l_sql & "pass_int_fallidos, pass_dias_log, pass_cambiar, pass_log_simult)"
		l_sql = l_sql & "VALUES ('" & l_pol_desc & "'," & l_pass_expira_dias & "," & l_pass_camb_dias & "," & l_pass_longitud & ","
		l_sql = l_sql & l_pass_historia & "," & l_pass_int_fallidos & "," & l_pass_dias_log & "," & l_pass_cambiar & "," & l_pass_log_simult & ")"
		
	 else
		l_sql = 		"UPDATE pol_cuenta "
		l_sql = l_sql & "SET pol_desc = '" & l_pol_desc & "', "
		l_sql = l_sql & "pass_expira_dias = " & l_pass_expira_dias & ", "
		l_sql = l_sql & "pass_camb_dias = " & l_pass_camb_dias & ", "
		l_sql = l_sql & "pass_longitud = " & l_pass_longitud & ", "
		l_sql = l_sql & "pass_historia = " & l_pass_historia & ","
		l_sql = l_sql & "pass_int_fallidos = " & l_pass_int_fallidos & ","
		l_sql = l_sql & "pass_dias_log = " & l_pass_dias_log & ","
		l_sql = l_sql & "pass_cambiar = " & l_pass_cambiar & ","
		l_sql = l_sql & "pass_log_simult = " & l_pass_log_simult & " "
		l_sql = l_sql & "WHERE pol_nro = " & l_pol_nro
		
	 end if
	l_cm.activeconnection = Cn
	cmExecute l_cm, l_sql, 0
		
	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
 end if
 l_rs.close
 cn.close
 
 Set l_rs = Nothing
 Set cn = Nothing
 Set l_cm = Nothing
 
%>
