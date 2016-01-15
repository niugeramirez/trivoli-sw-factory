<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!-- -----------------------------------------------------------------------------
Archivo     : verifica_campos.asp
Descripcion : Ejecuta una sql. Si no encuentra datos llama a Verif_sql_Valido(), sino Verif_sql_Invalido()
Creador		: Fernando Favre
Fecha 		: 16-12-2004
Modificado	:
------------------------------------------------------------------------------ -->
<%
on error goto 0
 Dim l_rs
 Dim l_sql
 Dim l_funcValida
 Dim l_funcNoValida
 
 l_sql 			= request("verif_sql")
 l_funcValida 	= request("funcValida")
 l_funcNoValida = request("funcNoValida")
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 rsOpen l_rs, cn, l_sql, 0
 
 if l_rs.eof then
 	response.write "<script>" & l_funcValida & ";</script>"
 else
 	response.write "<script>" & l_funcNoValida & ";</script>"
 end if
 l_rs.close
 
 Set l_rs = nothing
%>
