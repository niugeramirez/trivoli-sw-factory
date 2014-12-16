<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<% 
'Archivo: permisos_seg_01.asp
'Descripción: Asignacion de modelos de liquidacion para un concepto
'Autor : Alvaro Bayon
'Fecha: 05/11/2003
'Modificado:

'ADO
 Dim l_cm
 Dim l_sql
 Dim l_rs
 
' variables
 Dim l_grabar
 Dim l_menunro
  
'uso local
 dim l_iduser
 
 dim i
 dim l_lista
 Dim l_planta
 
' parámetros de entrada
 'La lista de registros a insertar viene en un string separados por comas
 l_grabar   = Request.QueryString("grabar")
 'Registro de la tabla entidad
 l_iduser = request.QueryString("iduser")
 
 'Al operar sobre varias tablas debo iniciar una transacción
 cn.BeginTrans
 
 'Creo los objetos ADO
 set l_rs = Server.CreateObject("ADODB.RecordSet")
 set l_cm = Server.CreateObject("ADODB.Command")
 
 'Busco en qué planta estoy
 l_sql = "SELECT tkt_planta.planro FROM tkt_planta"
 l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_lugar.planro = tkt_planta.planro"
 l_sql = l_sql & " INNER JOIN tkt_config ON tkt_config.lugnro = tkt_lugar.lugnro"
 rsOpen l_rs, cn, l_sql, 0 
 if not l_rs.eof then
	l_planta = l_rs("planro")
 end if
 l_rs.close
 
 'De la tabla relación busco los relacionados con la entidad y los borro
 l_sql = "SELECT menunro FROM tkt_usu_men_pla "
 l_sql = l_sql & " WHERE iduser = '" & l_iduser & "'"
 rsOpenCursor l_rs, cn, l_sql,0, adOpenKeyset
 l_cm.activeconnection = Cn
 do while not l_rs.eof
	l_sql = "DELETE FROM tkt_usu_men_pla "
	l_sql = l_sql & " WHERE menunro = " & l_rs("menunro")& " AND iduser = '" & l_iduser & "'"
	l_cm.CommandText = l_sql
	'response.write l_sql & vbcrlf & "<br>"
	cmExecute l_cm, l_sql, 0
 	l_rs.MoveNext
 loop
 l_rs.close
 
 'Separo los registros a insertar en un arreglo
 l_lista= Split(l_grabar,",")		
 i = 1
 do while i <= UBound(l_lista)-1
	l_menunro   = l_lista(i)
	l_sql = "INSERT INTO tkt_usu_men_pla "
	l_sql = l_sql & "(iduser, menunro,planro) "
	l_sql = l_sql & " VALUES ('" & l_iduser & "'," & l_menunro & "," & l_planta & ")"
	l_cm.CommandText = l_sql
	'response.write l_sql
	cmExecute l_cm, l_sql, 0
	i = i + 1
 loop	
 
 
 cn.CommitTrans
 'Fin de las operaciones sobre tablas
 
 cn.close
 Set cn = Nothing
 Set l_cm = Nothing
 
 response.write "<script>alert('Operación Realizada.');window.close();</script>"
 
%>
