<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo    : gen_lista_emp_00.asp
Descripción: Carga una 
Autor      : Scarpa D.
Fecha      : 26/11/2003
Modificado: 
-->
<% 
on error goto 0

Dim l_rs
Dim l_sql
Dim l_objSrc
Dim l_funcFin
Dim l_lista

l_sql     = request("sql")
l_objSrc  = request("objsrc")
l_funcFin = request("funcfin")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
rsOpen l_rs, cn, l_sql, 0 

l_lista = ""

do until l_rs.eof
    if l_lista = "" then
	   l_lista = l_rs(0)
	else   
	   l_lista = l_lista & "," & l_rs(0)
	end if
	
	l_rs.MoveNext
loop

l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
<script>
<%= l_objSrc & ".value = '" & l_lista & "'"  & vbCrLf %>

<%if l_funcFin <> "" then%>
<%= l_funcFin & "();"  & vbCrLf %>
<%end if%>

window.close();
</script>
</body>
</html>
