<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 


on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs

dim l_idvisita
dim l_idpracticarealizada
dim l_motivo
dim l_cantturnossimult  
dim l_cantsobreturnos     
dim l_turnoid

dim aux

function borrar_practica_realizada(idpracticarealizada,l_cm,Cn)

	dim l_sql


	l_sql = "DELETE FROM pagos "
	l_sql = l_sql & " WHERE idpracticarealizada = " & idpracticarealizada
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

	l_sql = "DELETE FROM practicasrealizadas "
	l_sql = l_sql & " WHERE id = " & idpracticarealizada
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	borrar_practica_realizada = 0
end function 

l_idvisita                 = request.Form("idvisita")
l_idpracticarealizada      = request.Form("idpracticarealizada")

'Al operar sobre varias tablas debo iniciar una transacción
cn.BeginTrans

	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	set l_cm = Server.CreateObject("ADODB.Command")
	
	l_cm.activeconnection = Cn

	'Si vino el ID de practica realizada por parametro es el caso del baja de practicas, entonces hago delete de la practca realizadaa y sus pagos sino,
	'hago delete de todas las practicas realizads y sus pagos, relacionados a la vista, luego hago delete de la visita
	if l_idpracticarealizada = "" or l_idpracticarealizada = 0 then

		l_sql = "select * "
		l_sql = l_sql & " from practicasrealizadas "
		l_sql = l_sql & " where idvisita = " & l_idvisita
		rsopencursor l_rs, cn, l_sql, 1, 1
		do while not l_rs.eof 

			aux = borrar_practica_realizada(l_rs("id"),l_cm,cn)
			'response.write "practica id "&l_rs("id")
			l_rs.movenext
		loop
		l_rs.close


		l_sql = "delete from visitas "
		l_sql = l_sql & " where id = " & l_idvisita
		l_cm.activeconnection = cn
		l_cm.commandtext = l_sql
		cmexecute l_cm, l_sql, 0
	else
		aux = borrar_practica_realizada(l_idpracticarealizada,l_cm,cn)
	end if
	
	Set l_cm = Nothing
cn.CommitTrans

Response.write "OK"
%>

