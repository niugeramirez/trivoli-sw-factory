<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<!--#include virtual="/turnos/shared/inc/sqls.inc"-->
<% 

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs

Dim l_id

Dim l_calfec
Dim l_descripcion

Dim l_idrecursoreservable

Dim l_pacienteid
Dim l_practicaid
Dim l_solicitadopor
Dim l_precio
Dim l_idvisita
Dim l_osid
Dim l_flag_particular
Dim l_idpracticarealizada
Dim l_practicarealizada

Dim l_idmediodepago
Dim l_idobrasocial
Dim l_nro
Dim l_importe
dim l_visitaid
dim l_practicasrealizadasid


'l_tipo 		           = request.querystring("tipo")
l_idrecursoreservable  = request.Form("idrecursoreservable")
l_pacienteid     	   = request.Form("pacienteid")
l_calfec               = request("calfec")
l_practicaid           = request("practicaid")
l_solicitadopor        = request("idmedicoderivador")

if l_solicitadopor = "" then
	l_solicitadopor = 0
end if

l_precio 			   = request.Form("precio2")
l_osid 				   = request.Form("osid")

l_idmediodepago        = request.Form("idmediodepago")
l_idobrasocial         = request.Form("idobrasocial")
l_nro                  = request.Form("nro")
l_importe              = request.Form("importe2")
l_idvisita 			   = request.Form("visitaid")
l_practicasrealizadasid = request.Form("practicasrealizadasid")

'response.write "l_idvisita "&l_idvisita

if l_importe = "" then l_importe = 0 end if

if l_idobrasocial = "" then l_idobrasocial = 0 end if
 


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

function BuscarmediopagoOS( )
	Dim l_rs
	Dim l_sql
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT * "
	l_sql = l_sql & " FROM mediosdepago "
	l_sql = l_sql & " WHERE flag_obrasocial = -1 " 
	l_sql = l_sql & " and empnro = " & Session("empnro")
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		BuscarmediopagoOS = l_rs("id")
	else
		BuscarmediopagoOS = 0
	end if
	l_rs.Close
	Set l_rs = Nothing
end function

'Al operar sobre varias tablas debo iniciar una transacción
cn.BeginTrans

Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")

l_sql = "SELECT isnull(obrassociales.flag_particular,0) flag_particular "
l_sql = l_sql & " FROM obrassociales "
l_sql  = l_sql  & " WHERE id = " & l_osid
l_sql = l_sql & " and empnro = " & Session("empnro")
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_flag_particular = l_rs("flag_particular")
end if
l_rs.Close

	'Si vino el ID de visita por parametro es el caso del alta de practicas, es decir ya esta dada de alta la visita
	if l_idvisita = "" or l_idvisita = 0 then
		l_sql = "INSERT INTO visitas "
		l_sql = l_sql & "(fecha, idrecursoreservable, idpaciente , idturno ,created_by,creation_date,last_updated_by,last_update_date,empnro) "
		l_sql = l_sql & "VALUES (" & cambiafecha(l_calfec,"YMD",true) & "," & l_idrecursoreservable  & "," &  l_pacienteid & ",0"&",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE(),'"& session("empnro") &"')"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		l_idvisita = codigogenerado("visitas")		
	end if
	
	'Si vino el ID de practica realizada por parametro es el caso del modificacion de practicas, entonces hago update y no doy de alta informacion de pagos
	if l_practicasrealizadasid = "" or l_practicasrealizadasid = 0 then
	
		l_sql = "INSERT INTO practicasrealizadas (idvisita , idpractica , idsolicitadapor , precio ,created_by,creation_date,last_updated_by,last_update_date, empnro ) "
		l_sql = l_sql & " VALUES ( " & l_idvisita & ","  & l_practicaid & "," & l_solicitadopor & "," & l_precio &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE(),'"& session("empnro") &"')"	

		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0	
		
		' Si tiene Obra Social registro el Pago (solo si tiene precio, para no generar informacion innecesaria)
		'if l_flag_particular = 0 and l_precio <> 0 then
		if l_importe <> 0 then
			l_practicarealizada = codigogenerado("practicasrealizadas")				
			
			l_sql = "INSERT INTO pagos "
			l_sql = l_sql & "( idpracticarealizada, idmediodepago, idobrasocial, nro, fecha , importe ,created_by,creation_date,last_updated_by,last_update_date, empnro) "
			'l_sql = l_sql & "VALUES (" & l_practicarealizada  & "," & BuscarmediopagoOS( ) & "," & l_osid & "," & cambiafecha(l_calfec,"YMD",true) & "," & l_precio &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE(),'"& session("empnro") &"')"
			l_sql = l_sql & "VALUES (" & l_practicarealizada  & "," & l_idmediodepago & "," & l_idobrasocial & ",'" & l_nro & "'," & cambiafecha(l_calfec,"YMD",true) & "," & l_importe &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE(),'"& session("empnro") &"')"		
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0 
		
		end if	
	else
		l_sql = "UPDATE practicasrealizadas"
		l_sql = l_sql & " SET idpractica = " & l_practicaid
		l_sql = l_sql & " , idsolicitadapor = " & l_solicitadopor		
		l_sql = l_sql & " , precio = " & l_precio
		l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
		l_sql = l_sql & "    ,last_update_date = GETDATE()" 		
  	    l_sql = l_sql & " WHERE id = " & l_practicasrealizadasid

		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0		
	end if

		
Set l_cm = Nothing
cn.CommitTrans

Response.write "OK"

%>