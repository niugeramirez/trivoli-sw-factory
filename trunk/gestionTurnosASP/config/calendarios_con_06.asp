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
Dim l_descripcion

Dim l_calfec
Dim l_calhordes
Dim l_calhordes1
Dim l_calhordes2
Dim l_calhorhas
Dim l_calhorhas1
Dim l_calhorhas2
Dim l_intervaloTurnoMinutos

Dim l_calhorhasmasint

Dim texto
Dim texto1
Dim texto2

texto = ""
l_tipo		    = request.QueryString("tipo")
l_id            = request.QueryString("id")
l_descripcion 	= request.QueryString("descripcion")
l_calfec    	= request.QueryString("calfec")
l_calhordes1 	= request.QueryString("calhordes1")
l_calhordes2 	= request.QueryString("calhordes2")
l_calhorhas1 	= request.QueryString("calhorhas1")
l_calhorhas2 	= request.QueryString("calhorhas2")
l_intervaloTurnoMinutos = request.QueryString("intervaloTurnoMinutos")

l_calhordes = l_calhordes1 & ":" & l_calhordes2 & ":00"
l_calhorhas = l_calhorhas1 & ":" & l_calhorhas2 & ":00"
l_calhorhasmasint = DateAdd("n", cint(l_intervaloTurnoMinutos), l_calhorhas)

response.write "l_calhorhas " & l_calhorhas & "<br>"
response.write "l_intervaloTurnoMinutos " & l_intervaloTurnoMinutos & "<br>"
response.write "hora sumada " & l_calhorhasmasint & "<br>"
'response.end


'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico si existen Calendarios que tienen mayor fecha de inicio que la hora Desde Ingresada 
l_sql = "SELECT *"
l_sql = l_sql & " FROM calendarios "
l_sql = l_sql & " WHERE idrecursoreservable=" & l_id
'l_sql = l_sql & " AND estado='ACTIVO'"
l_sql = l_sql & " AND CONVERT(VARCHAR(10), calendarios.fechahorafin, 101)  = " & cambiafecha (l_calfec,  "mm/dd/yyyy","" )  
l_sql = l_sql & " AND calendarios.fechahorafin > " & cambiaformato (l_calfec,l_calhordes )  
l_sql = l_sql & " order by fechahorafin desc "
'response.write l_sql

rsOpen l_rs, cn, l_sql, 0
'response.write l_rs.eof 
'response.end
if not l_rs.eof then
    texto1 =  "Existen Calendarios con fecha de Finalizacion posterior a la fecha de Inicio del Dia que se quiere ingresar."
end if 
l_rs.close

'Verifico si existen Calendarios que tienen mayor fecha de inicio que la hora Desde Ingresada 
l_sql = "SELECT *"
l_sql = l_sql & " FROM calendarios "
l_sql = l_sql & " WHERE idrecursoreservable=" & l_id
'l_sql = l_sql & " AND estado='ACTIVO'"
l_sql = l_sql & " AND CONVERT(VARCHAR(10), calendarios.fechahorafin, 101)  = " & cambiafecha (l_calfec,  "mm/dd/yyyy","" )  
l_sql = l_sql & " AND calendarios.fechahorainicio < " & cambiaformato (l_calfec,l_calhorhasmasint )  
l_sql = l_sql & " order by fechahorafin asc "
'response.write l_sql

rsOpen l_rs, cn, l_sql, 0
'response.write l_rs.eof 
'response.end
if not l_rs.eof then
    texto2 =  "Existen Calendarios con fecha/hora de Inicio anterior a la fecha/hora de Finalizacion del Dia mas el Intervalo de Tiempo del Turno que se quiere ingresar."
end if 
l_rs.close

%>

<script>
<% 
 if texto1 <> "" and texto2 <> ""  then
 	texto = texto2 & texto1
%>
   parent.invalido('<%= texto %> ')
<% else%>
   parent.valido();
<% end if%>
</script>

<%
Set l_rs = Nothing
%>

