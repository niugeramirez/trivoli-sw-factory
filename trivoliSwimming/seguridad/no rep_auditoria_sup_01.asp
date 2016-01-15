<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/sqls.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/util.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/adovbs.inc"-->
<!--
Archivo        : rep_auditoria_sup_01.asp
Descripción    : Reporte Auditoría - Generar Proceso
Autor          : JMH
Fecha Creacion : 20/01/2005
Modificado     : 
			21/07/2005 - Fapitalle N. - Agregar filtro por empleado
-->

<% 
on error goto 0

Const l_Max_Lineas_X_Pag = 50
const l_nro_col = 4

Dim l_arrTmp 
Dim l_arrTmp2
Dim l_tituloR 

Dim l_rs
Dim l_rs2
Dim l_rs3
Dim l_rs4

Dim l_cm

Dim l_sql
Dim l_sqlemp

Dim l_nrolinea
Dim l_pagina
Dim l_totalemp

Dim l_encabezado
Dim l_corte
Dim l_cambioEmp
Dim l_conc_detdom
Dim l_listaproc
Dim l_procant
Dim l_terant
Dim l_pliqdesc
Dim l_pliqant
Dim l_pliqmesant
Dim l_pliqanioant

Dim l_total_mbasico 
Dim l_total_mneto   
Dim l_total_mmsr    
Dim l_total_masi_flia
Dim l_total_mDtos
Dim l_total_mbruto

'Parametros
 Dim l_filtro ' Viene el filtro comun: empest, legajo, 
 Dim l_orden
 Dim l_taccion
 Dim l_accion
 Dim l_tusuario
 Dim l_usuario
 dim l_caudnro
 Dim l_fechadesde
 Dim l_fechahasta
 Dim l_campos
 Dim l_emptipo
 Dim l_empleados
 Dim l_list_emp
 Dim l_params
 Dim l_titulofiltro ' Viene el titulo armado segun filtro

l_titulofiltro	= request("tfiltro")
l_filtro 		= request("filtro")
l_orden         = request("orden")

l_taccion		= request.querystring("acciones")
l_accion		= request.querystring("acnro")
l_tusuario 		= request.querystring("usuarios")
l_usuario 		= request.querystring("iduser")
l_caudnro 		= request.querystring("caudnro")
l_fechadesde	= request.querystring("fechadesde")
l_fechahasta	= request.querystring("fechahasta")
l_campos   	    = request.form("campos")

l_emptipo		= request.querystring("emptipo")
l_empleados		= request.form("lista")

l_list_emp		= ""
l_params		= ""

' debug de pasaje de parametros por form
'response.write request.form("lista")  & "<br>"
'response.write request.form("campos")  & "<br>"
'response.write l_filtro  & "<br>"
'response.end

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/rhprox2/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%
Set l_rs  = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
Set l_rs3 = Server.CreateObject("ADODB.RecordSet")
Set l_rs4 = Server.CreateObject("ADODB.RecordSet")

if CInt(l_emptipo) = 2  then
   l_sql = "SELECT ternro FROM v_empleado WHERE empleg =" & l_empleados
   
   rsOpen l_rs, cn, l_sql, 0	   
   
   if not l_rs.eof then
      l_empleados = l_rs("ternro")
   end if
   
   l_rs.close

end if

' ------------------------------------------------------------------------------------------------------------------
' generarSQLProc(tipoPorc,desde,hasta) :
' parametros:
'    * tipoPorc     : tipo de proceso a insertar
'    * desde, hasta : fechas desde y hasta sin formatear
'    * parametros   : parametros opcionales del procedimiento
' ------------------------------------------------------------------------------------------------------------------
function generarSQLProc(tipoPorc,parametros,fechadesde,fechahasta)

Dim l_id
Dim l_hora
Dim l_dia
Dim l_sqlp

l_id   = Session("Username")
l_hora = mid(time,1,8)
l_dia  = cambiafecha(Date,"YMD",true)

l_sqlp =          " INSERT INTO batch_proceso "
l_sqlp = l_sqlp & " (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
l_sqlp = l_sqlp & " bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados,bprcurgente) "
l_sqlp = l_sqlp & " VALUES (" & tipoPorc & "," & l_dia & ", '"& l_id &"','"& l_hora &"' "
l_sqlp = l_sqlp & " ," & cambiafecha(fechadesde,"YMD",true) & ", " & cambiafecha(fechahasta,"YMD",true) 
l_sqlp = l_sqlp & " , '" & parametros & "', 'Pendiente', null , null, null, null, 0, null,0)"

generarSQLProc = l_sqlp

end function 'generarSQLProc()

' ------------------------------------------------------------------------------------------------------------------
' insertarempleados(codproc) :
' ------------------------------------------------------------------------------------------------------------------
function insertarempleados(codproc, listaemp)
	Dim l_sqli
	Dim l_sqls
	Dim i
	Dim l_ter

	i = 0
	For Each l_ter In listaemp 'busco el ternro de cada empleado
		
		l_sqli = "INSERT INTO batch_empleado "
		l_sqli = l_sqli & " (bpronro, ternro, estado) "
		l_sqli = l_sqli & " VALUES (" & codproc & "," & l_ter & ",'" & i & "')"
		l_cm.CommandText = l_sqli
		cmExecute l_cm, l_sqli, 0

		i = i + 1

	Next

end function 'insertarempleados(codproc, listaempleados)

' ------------------------------------------------------------------------------------------------------------------
' codprocbatch() :
' ------------------------------------------------------------------------------------------------------------------
function codprocbatch()
	l_sql = fsql_seqvalue("next_id","batch_proceso")
	rsOpenCursor l_rs, cn, l_sql, 0, adOpenKeyset
	codprocbatch=l_rs("next_id")
	l_rs.Close
end function 'codprocbatch()

' ------------------------------------------------------------------------------------------------------------------
' valor() :
' ------------------------------------------------------------------------------------------------------------------
function valor(v)
  if v <> "" and v <> "0" then
    valor = v
  else
    valor = "0"
  end if
end function 'valor()

Dim l_codproc
Dim l_arrempl

set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = cn

' ------------------------------------------------------------------------------------------------------------------
' EMPIEZA TRANSACCION
' ------------------------------------------------------------------------------------------------------------------
cn.BeginTrans

'Ingreso el proceso a la tabla, parametros lleva el numero de legajo del empleado
l_params = l_taccion & "@" & l_accion & "@" & l_tusuario & "@" & l_usuario & "@" & l_caudnro & "@" & l_campos & "@" & l_emptipo
l_sql = generarSQLProc(70,l_params,l_fechadesde, l_fechahasta)
cmExecute l_cm, l_sql, 0

'Ingreso la lista de empleados a la tabla
l_codproc = codprocbatch()

if CInt(l_emptipo) = 2 OR CInt(l_emptipo) = 1 then
   l_arrempl = split(l_empleados,",")
   insertarempleados l_codproc, l_arrempl
end if

cn.CommitTrans
' ------------------------------------------------------------------------------------------------------------------
' TERMINA TRANSACCION
' ------------------------------------------------------------------------------------------------------------------

set l_rs = Nothing

cn.Close
set cn = Nothing
%>
</table>
<script>
   parent.document.ifrm.location = 'rep_auditoria_sup_02.asp?bpronro=<%= l_codproc%>';
</script>
</body>
</html>


