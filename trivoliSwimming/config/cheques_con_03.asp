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
dim l_numero
dim l_fecha_emision
dim l_fecha_vencimiento  
dim l_idbanco  
dim l_importe
dim l_flag_propio
dim l_flag_emitidopor_cliente
dim l_emisor
dim l_validacion_bcra
dim l_flag_cobrado_pagado


l_tipo 		               = request.Form("tipo")
l_id                       = request.Form("id")
l_numero                   = request.Form("numero")
l_fecha_emision	           = request.Form("fecha_emision")
l_fecha_vencimiento        = request.Form("fecha_vencimiento")
l_idbanco  		           = request.Form("idbanco")
l_importe		           = request.Form("importe2")
l_flag_propio			   = request.Form("flag_propio")
l_flag_emitidopor_cliente			   = request.Form("flag_emitidopor_cliente")
l_emisor			       = request.Form("emisor")
l_validacion_bcra			= request.Form("validacion_bcra")
l_flag_cobrado_pagado 		= request.Form("flag_cobrado_pagado")

'inicializo los campos que pueden venir en nulo
if len(l_fecha_emision) = 0 then
	l_fecha_emision = "null"
else 
	l_fecha_emision = cambiafecha(l_fecha_emision,"YMD",true)	
end if 

if len(l_fecha_vencimiento) = 0 then
	l_fecha_vencimiento = "null"
else 
	l_fecha_vencimiento = cambiafecha(l_fecha_vencimiento,"YMD",true)	
end if 
'fin inicializacion campos en nulo


'response.write "l_tipo"&l_tipo & "<br>"
'response.write "l_id"&l_id & "<br>"
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
	 l_sql = "INSERT INTO cheques  "
	 l_sql = l_sql & " (numero, fecha_emision, fecha_vencimiento, id_banco , importe, flag_propio, emisor,flag_emitidopor_cliente,  empnro, created_by,creation_date,last_updated_by,last_update_date,validacion_bcra)"
	 l_sql = l_sql & " VALUES ('" & l_numero & "'," & l_fecha_emision & "," & l_fecha_vencimiento & "," & l_idbanco & "," & l_importe & "," & l_flag_propio & ",'" & l_emisor &"',"&  l_flag_emitidopor_cliente  &",'" & session("empnro")& "','"   &session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE(),'"&l_validacion_bcra&"')"
	
 else
	l_sql = "UPDATE cheques "
	l_sql = l_sql & " SET numero    = '" & l_numero & "'"
	l_sql = l_sql & "    ,fecha_emision  = " & l_fecha_emision & ""	
	l_sql = l_sql & "    ,fecha_vencimiento  = " & l_fecha_vencimiento & ""	
	l_sql = l_sql & "    ,id_banco   = " & l_idbanco & ""
	l_sql = l_sql & "    ,importe      = " & l_importe & ""
	l_sql = l_sql & "    ,flag_propio = " & l_flag_propio & ""
	l_sql = l_sql & "    ,flag_emitidopor_cliente = " & l_flag_emitidopor_cliente & ""
	l_sql = l_sql & "    ,emisor  =  '" & l_emisor & "'"	
	l_sql = l_sql & "    ,validacion_bcra  =  '" & l_validacion_bcra & "'"		
	l_sql = l_sql & "    ,flag_cobrado_pagado = " & l_flag_cobrado_pagado & ""		
	l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
	l_sql = l_sql & "    ,last_update_date = GETDATE()" 	
	l_sql = l_sql & " WHERE id = " & l_id
 end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

if l_tipo = "A" then
	l_id = codigogenerado("cheques")
end if

Set l_cm = Nothing

cn.CommitTrans 

'Response.write "OK"
Response.write "[{""resultado"":""OK"",""id"":""" & l_id & """}]"
%>

