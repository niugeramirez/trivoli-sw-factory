<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->

<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<%
'---------------------------------------------------------------------------------
'Archivo	: relacionar_empleados_eva_02.asp
'Descripción: Crea los registros de tablas Batch para que se ejecute el proceso
'			de creacion de Formularios de Evaluacion
'Autor		: CCRossi
'Fecha		: 18-01-2005
'Modificado	: 
'----------------------------------------------------------------------------------
  
' variables
  Dim l_cm
  Dim l_sql
  Dim l_rs
  
' locales
  dim l_codproc
  dim i
  dim arr1
  dim arr2
  dim l_lista1 
  dim l_lista2
  dim l_totalemp  
  dim l_parametros  
  
' parametros de entrada
  Dim l_listempleados
  Dim l_evaevenro 
  
  l_listempleados  = Request("listempleados")
  l_evaevenro	   = request("evaevenro")

' ------------------------------------------------------------------------------------------------------------------
' generarSQLProc(tipoPorc,desde,hasta) :
' parametros:
'    * tipoProc     : tipo de proceso a insertar
'    * desde, hasta : fechas desde y hasta sin formatear
'    * parametros   : parametros opcionales del procedimiento
' ------------------------------------------------------------------------------------------------------------------
function generarSQLProc(tipoPorc,parametros)

Dim l_id
Dim l_hora
Dim l_dia
Dim l_sqlp

l_id   = Session("Username")
l_hora = mid(time,1,8)
l_dia  = cambiafecha(Date,"YMD",true)

l_sqlp =          " INSERT INTO batch_proceso (btprcnro, bprcfecha, iduser, bprchora, bprcparam, bprcestado, empnro, bprcurgente) "
l_sqlp = l_sqlp & " VALUES (" & tipoPorc & "," & l_dia & ", '"& l_id &"','"& l_hora &"' "
l_sqlp = l_sqlp & " , '" & parametros & "', 'Pendiente', 0, 0)"

generarSQLProc = l_sqlp

end function 'generarSQLProc()

' ------------------------------------------------------------------------------------------------------------------
' insertarempleados(codproc) :
' ------------------------------------------------------------------------------------------------------------------
function insertarempleados(codproc, listaempleados)
	Dim l_sqli
	Dim arregloemp 
	Dim i
	
	arregloemp = Split(listaempleados,",")
	
	l_totalemp = UBound(arregloemp)-1
	
	For i=0 To UBound(arregloemp)


'response.write ("<script>alert('"&arregloemp(i)&"')</script>")		
	    if (CInt(arregloemp(i)) <> 0) then
			l_sql = "select * from batch_empleado where bpronro =" & codproc &" and ternro="& arregloemp(i)
			rsOpen l_rs, cn, l_sql, 0
			if l_rs.eof then
				l_sqli = "INSERT INTO batch_empleado (bpronro, ternro) VALUES (" & codproc & "," & arregloemp(i) & ")"
				l_cm.CommandText = l_sqli
				cmExecute l_cm, l_sqli, 0
			end if
			l_rs.Close	
	    end if
	Next 

end function 'insertarempleados(codproc, listaempleados)

' ------------------------------------------------------------------------------------------------------------------
' codprocbatch() :
' ------------------------------------------------------------------------------------------------------------------
function codprocbatch()
	l_sql = fsql_seqvalue("next_id","batch_proceso")
	rsOpen l_rs, cn, l_sql, 0
	codprocbatch=l_rs("next_id")
	l_rs.Close
end function 'codprocbatch()

'=============================================================================
'								BODY
'=============================================================================

'convertir las listas a unicamente ternros
l_lista1 = "0"
arr1 = Split(l_listempleados,",")
i=0
do while i<=Ubound(arr1) 
  
    arr2 = split(arr1(i),"@")
    l_lista1 = l_lista1 & "," & arr2(0)
	
i = i+1
loop	

l_listempleados = l_lista1

'Preparo el string del parametro - El formato es (evaevenro.listinicial)
l_parametros = l_evaevenro 

Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Comienzo creacion de registros BATCH ===========================================
set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = cn

cn.BeginTrans


'Ingreso el proceso a la tabla con codigo 69=creacion Formularios Eva
l_sql = generarSQLProc(69,l_parametros)
	
cmExecute l_cm, l_sql, 0
	
'Ingreso la lista de empleados a la tabla
l_codproc = codprocbatch()
l_totalemp = 0

'response.write ("<script>alert('"&l_listempleados&"')</script>")

insertarempleados l_codproc, l_listempleados
	

cn.CommitTrans

%>
<script>
   //parent.document.ifrm.location = 'relacionar_empleados_eva_03.asp?bpronro=<%= l_codproc%>';
    abrirVentana('relacionar_empleados_eva_03.asp?bpronro=<%= l_codproc%>&totalemp=<%= l_totalemp%>','',320,340);
</script>


