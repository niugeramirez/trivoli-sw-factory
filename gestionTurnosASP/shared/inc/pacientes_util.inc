
<%
function get_proximo_histnum(empnro,l_cm)
	Dim l_rs
	Dim l_sql
	dim l_ultimo_histnum
	
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	
	l_sql = "select ISNULL(empresa.ultimo_histnum,0) ultimo_histnum"
	l_sql  = l_sql  & " FROM empresa "	
	l_sql = l_sql & " WHERE empresa.id = " & empnro
	rsOpen l_rs, cn, l_sql, 0 
	
	if not l_rs.eof then
		l_ultimo_histnum = l_rs("ultimo_histnum")+1	
	end if
	l_rs.Close
	
	'Si la empresa no tiene configurado el maximo numero de historia clinica, busco el maximo cargado en la tabla de pacientes
	if l_ultimo_histnum = 1 then
		l_sql = "select ISNULL(max(CAST(nrohistoriaclinica as int)),0)  nro_max "
		l_sql  = l_sql  & " FROM clientespacientes "
		l_sql = l_sql & " where ISNUMERIC(nrohistoriaclinica) <> 0 "
		l_sql = l_sql & " and ISNULL(nrohistoriaclinica,'0') <> '0' "
		l_sql = l_sql & " AND clientespacientes.empnro = " & empnro	
		
		rsOpen l_rs, cn, l_sql, 0 
		
		if not l_rs.eof then
			l_ultimo_histnum = l_rs("nro_max")	+1
		end if
		l_rs.Close	
	end if
	
	l_sql = "update empresa "
	l_sql = l_sql & " set empresa.ultimo_histnum =  "&l_ultimo_histnum
	l_sql = l_sql & " where empresa.id =  " &session("empnro")
	
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	get_proximo_histnum = l_ultimo_histnum
end function

function check_genera_histnum(empnro)
	Dim l_rs
	Dim l_sql
	dim l_flag_genera_histnum
	
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	
	l_sql = "select ISNULL(empresa.flag_genera_histnum,0) flag_genera_histnum"
	l_sql  = l_sql  & " FROM empresa "	
	l_sql = l_sql & " WHERE empresa.id = " & empnro
	rsOpen l_rs, cn, l_sql, 0 
	
	if not l_rs.eof then
		l_flag_genera_histnum = l_rs("flag_genera_histnum")
	else
		l_flag_genera_histnum = "0"
	end if
	l_rs.Close
	
	if l_flag_genera_histnum = "0" then
		check_genera_histnum = false
	else
		check_genera_histnum = true
	end if
end function
%>
