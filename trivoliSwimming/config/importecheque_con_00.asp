<% Option Explicit %>

<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->

<% 

Dim l_id
Dim l_importe

Dim l_rs
Dim l_sql

l_id = request("id")



'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

	if l_id = "" then
		l_importe = 0
	else
		
		l_sql = "SELECT   cheques.importe"
		l_sql  = l_sql  & " FROM cheques "
		l_sql = l_sql & " where cheques.empnro = " & Session("empnro") 
		l_sql = l_sql & " and  cheques.id = " & l_id 
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then
			l_importe = l_rs("importe")
		else
			l_importe = 0
		end if
		l_rs.Close
		
	end if
	
%>
<script>
	parent.actualizarimporte('<%= l_importe %>')
</script>
