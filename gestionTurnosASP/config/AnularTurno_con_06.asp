<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 



Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_id
Dim l_dni
Dim l_nrohistoriaclinica

Dim l_hd
Dim l_md
Dim l_hh
Dim l_mh

Dim l_opc

Dim l_horadesde
Dim l_horahasta
Dim l_fechadesde
Dim l_fechahasta

Dim l_idrecursoreservable

Dim texto

texto = ""
l_tipo		    	 = request.QueryString("tipo")
l_id            	 = request.QueryString("id")
l_fechadesde         = request.QueryString("qfechadesde")
l_fechahasta       	 = request.QueryString("qfechahasta")
l_opc 				 = request.QueryString("opc")
l_hd			     = request.QueryString("hd") 
l_md			     = request.QueryString("md")
l_hh			     = request.QueryString("hh")
l_mh			     = request.QueryString("mh")
l_idrecursoreservable = request.QueryString("idrecursoreservable")

'response.write "ra" & l_opc & l_idrecursoreservable

l_horadesde = l_hd & ":" & l_md
l_horahasta = l_hh & ":" & l_mh

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")


'Response.write "<script>alert('l_opc  " &l_opc&" .');</script>"
'Response.write "<script>alert('l_tipo  " &l_tipo&" .');</script>"

if l_opc = 1 then
	texto = ""
	'Verifico que no tenga calendarios asignados
								
	l_sql = "SELECT * , calendarios.id "
	l_sql = l_sql & " FROM calendarios "
	l_sql = l_sql & " WHERE id = " & l_id
	
	'if l_tipo = "B" then ' Bloquear
		l_sql = l_sql & " AND estado='ACTIVO'"
		l_sql = l_sql & " AND calendarios.id in (SELECT idcalendario FROM turnos )" 
	'else
	'	l_sql = l_sql & " AND estado='ACTIVO'"
	'end if
	l_sql = l_sql & " AND idrecursoreservable=" & l_idrecursoreservable
	l_sql = l_sql & " and calendarios.empnro = " & Session("empnro")  
								
	'if l_tipo = "M" then
	'	l_sql = l_sql & " AND id <> " & l_id
	'end if
	'l_sql = l_sql & " and clientespacientes.empnro = " & Session("empnro")   
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
	    texto =  "Tiene Calendarios asignados para el Rango de Fechas ingresado." '& l_rs("id") 
	end if 
	l_rs.close	
	
else

	'Verifico que no tenga calendarios asignados
								
	l_sql = "SELECT * , calendarios.id "
	l_sql = l_sql & " FROM calendarios "
	l_sql = l_sql & " WHERE fechahorainicio >=" & cambiaformato (l_fechadesde,l_horadesde )
	l_sql = l_sql & " AND fechahorainicio<=" & cambiaformato (l_fechahasta,l_horahasta )
	
	'if l_tipo = "B" then ' Bloquear
		l_sql = l_sql & " AND estado='ACTIVO'"
		l_sql = l_sql & " AND calendarios.id in (SELECT idcalendario FROM turnos )" 
	'else
	'	l_sql = l_sql & " AND estado='ACTIVO'"
	'end if
	l_sql = l_sql & " AND idrecursoreservable=" & l_idrecursoreservable
	l_sql = l_sql & " and calendarios.empnro = " & Session("empnro")  
								
	'if l_tipo = "M" then
	'	l_sql = l_sql & " AND id <> " & l_id
	'end if
	'l_sql = l_sql & " and clientespacientes.empnro = " & Session("empnro")   
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
	    texto =  "Tiene Calendarios asignados para el Rango de Fechas ingresado." '& l_rs("id") 
	end if 
	l_rs.close
	
end if
%>

<script>
<% 
 if texto <> "" then
%>
   parent.invalido('<%= texto %>')
<% else%>
   parent.valido();
<% end if%>
</script>

<%
Set l_rs = Nothing
%>

