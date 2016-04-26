<% Option Explicit %>

<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->

<% 

Dim l_id
Dim l_flagcompra
Dim l_flagventa

Dim l_rs
Dim l_sql

l_id = request("id")



'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

	if l_id = "" then
		l_flagcompra = 0
		l_flagventa = 0
	else
		
		l_sql = "SELECT   isnull(tiposMovimientoCaja.flagcompra,0) flagcompra , isnull(tiposMovimientoCaja.flagventa,0) flagventa"
		l_sql  = l_sql  & " FROM tiposMovimientoCaja "
		l_sql = l_sql & " where tiposMovimientoCaja.empnro = " & Session("empnro") 
		l_sql = l_sql & " and  tiposMovimientoCaja.id = " & l_id 
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then
			l_flagcompra = l_rs("flagcompra")
			l_flagventa = l_rs("flagventa")
		
		else
			l_flagcompra = 0
			l_flagventa = 0
		end if
		l_rs.Close
		
	end if



	
%>
<script>
	parent.actualizarflag('<%= l_flagcompra%>', '<%= l_flagventa %> ')
</script>
