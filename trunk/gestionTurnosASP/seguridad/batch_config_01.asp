<%  Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
on error goto 0
' Declaracion de Variables locales  -------------------------------------
Dim l_tipo
Dim l_cm
Dim l_sql
dim l_rs

Dim l_btc_tmp_esp_no_resp
Dim l_btc_tmp_esp_sin_prog
Dim l_btc_tmp_lect_reg
Dim l_btc_tmp_dorm
Dim l_btc_usa_reg
Dim l_btc_max_proc
Dim l_btc_mult_logs
Dim l_btc_path_proc
Dim l_btc_path_logs
Dim l_btc_form_fecha

' traer valores del form de alta -------------------------------------

l_btc_tmp_esp_no_resp  = Request.form("tenr")
l_btc_tmp_esp_sin_prog = Request.form("tesp")
l_btc_tmp_lect_reg     = Request.form("tldr")
l_btc_tmp_dorm         = Request.form("tdd")
l_btc_usa_reg          = Request.form("ureg")
l_btc_max_proc         = Request.form("mproc")
l_btc_mult_logs        = Request.form("march")
l_btc_path_proc        = Request.form("pproc")
l_btc_path_logs        = Request.form("plogs")
l_btc_form_fecha       = Request.form("fecha")

if l_btc_tmp_lect_reg = "" then
   l_btc_tmp_lect_reg = 0
end if

if l_btc_usa_reg = "" then
   l_btc_usa_reg = 0
end if

if l_btc_mult_logs = "" then
   l_btc_mult_logs = 0
end if

l_btc_path_proc  = replace(l_btc_path_proc,"'","")
l_btc_path_logs  = replace(l_btc_path_logs,"'","")
l_btc_form_fecha = replace(l_btc_form_fecha,"'","")

' trasnformar valor de checkboxes en valores logicos --------------------------

set l_cm = Server.CreateObject("ADODB.Command")
'BORRO LA TABLA DE CONFIGURACION DE BATCH PROCESO
l_sql = "DELETE batch_config "
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

'INSERTO LA NUEVA CONFIGURACION

l_sql = "INSERT INTO batch_config "
l_sql = l_sql & "(btc_tmp_esp_no_resp, btc_tmp_esp_sin_prog, btc_tmp_lect_reg, btc_tmp_dorm, btc_usa_reg, "
l_sql = l_sql & " btc_max_proc, btc_mult_logs, btc_path_proc, btc_path_logs, btc_form_fecha ) "
l_sql = l_sql & " values (" 
l_sql = l_sql & l_btc_tmp_esp_no_resp  & ","
l_sql = l_sql & l_btc_tmp_esp_sin_prog & ","
l_sql = l_sql &	l_btc_tmp_lect_reg     & ","
l_sql = l_sql &	l_btc_tmp_dorm         & ","
l_sql = l_sql &	l_btc_usa_reg          & ","
l_sql = l_sql &	l_btc_max_proc         & ","
l_sql = l_sql &	l_btc_mult_logs        & ",'"
l_sql = l_sql &	l_btc_path_proc        & "','"
l_sql = l_sql &	l_btc_path_logs        & "','"
l_sql = l_sql &	l_btc_form_fecha       & "')"

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0	

Response.write "<script>alert('Operación Realizada.');window.close();</script>"
%>
