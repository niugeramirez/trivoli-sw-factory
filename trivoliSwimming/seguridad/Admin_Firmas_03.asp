<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<% 
Dim l_proximo
Dim l_anterior
Dim l_usuario
Dim l_actual
Dim l_tipo
Dim l_codigo
Dim l_esfin
Dim l_esModif
Dim l_esPrimero
Dim l_Descripcion
Dim l_PuedeVer
Dim l_secuencia
Dim l_sql
Dim l_sql1
Dim l_cm

l_tipo        = request("tipo")
l_codigo      = request("codigo")
l_usuario     = session("username") 
l_proximo     = request("codproximo") 
l_anterior    = request("codanterior") 
l_actual      = request("codactual") 
l_esfin       = request("esfin") 
l_esModif     = request("esmodif") 
l_esPrimero   = request("esprimero") 
l_Secuencia   = request("secuencia") 
l_Descripcion = request("descripcion") 

if l_esModif = "false" then  'debo crear uno 
   if l_esPrimero = "false" then  'y actualizar el anterior
 	  l_sql = "update cysfirmas "
	  l_sql = l_sql & "set cysfiryaaut = -1 where cystipnro = " & l_tipo & " and cysfircodext = '" & l_codigo & "' and cysfirsecuencia = " & l_secuencia
   end if 
   l_secuencia = l_secuencia + 1  'incremento la secuencia para el proximo
   if l_esfin = "true" then
     l_esfin = "-1"
   else
     l_esfin = "0"
   end if

   l_sql1 = "insert into cysfirmas "
   l_sql1 = l_sql1 & "(cysfirautoriza, cysfirfecaut, cysfirmhora, cysfirusuario, cysfirdestino, cystipnro, cysfircodext, cysfirsecuencia, cysfirdes, cysfirfin, cysfiryaaut) " 
   l_sql1 = l_sql1 & "values ('" & l_actual & "', " & cambiafecha(date, "YMD", true) & ", '" & time & "', '" & l_usuario & "', '" & l_proximo & "', " & l_tipo & ", '" & l_codigo & "', " & l_secuencia & ", '" & l_descripcion & "', " & l_esfin & ", " & l_esfin & ")"
else
	l_sql = "update cysfirmas "
	l_sql = l_sql & "set cysfirdestino = '" & l_proximo & "', cysfirfecaut = " & cambiafecha(date, "YMD", true) & ", cysfirmhora = '" & time & "' where cystipnro = " & l_tipo & " and cysfircodext = '" & l_codigo & "' and cysfirsecuencia = " & l_secuencia
end if

cn.begintrans
set l_cm = Server.CreateObject("ADODB.Command")

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
l_cm.CommandText = l_sql1
cmExecute l_cm, l_sql1, 0
cn.committrans
Set cn = Nothing
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'Admin_Firmas_01.asp';window.close();</script>"

%>
