<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: balanzas_con_03.asp
'Descripción: ABM de Balanzas
'Autor : Gustavo Manfrin
'Fecha: 27/04/2005

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_balnro
Dim l_baldes
Dim l_balcod
Dim l_balact
Dim l_planro
Dim l_balvpc
Dim	l_balmarca
Dim l_balconexion


l_tipo = request.querystring("tipo")
l_balnro = request.Form("balnro")
l_baldes = request.Form("baldes")
l_balcod = request.Form("balcod")
if request.Form("balact")="on" then  
   l_balact =-1 
else 
   l_balact =0 
end if
l_planro = request.Form("planro")
if request.Form("balvpc")="on" then 
   l_balvpc = -1 
else 
   l_balvpc = 0 
end if
l_balmarca = request.Form("balmarca")
l_balconexion = request.Form("balcon")


	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO tkt_balanza"
		l_sql = l_sql & " (balcod, baldes, balact, planro, balvpc, balmarca, balconexion)"
		l_sql = l_sql & " VALUES ('" 
		l_sql = l_sql & l_balcod & "','" & l_baldes & "','" & l_balact & "','"& l_planro 
    	l_sql = l_sql & "','" & l_balvpc & "','" & l_balmarca & "','" & l_balconexion
		l_sql = l_sql & "')"
	else
		l_sql = "UPDATE tkt_balanza"
		l_sql = l_sql & " SET baldes = '" & l_baldes & "'"
		l_sql = l_sql & ", balcod = '" & l_balcod & "'"
		l_sql = l_sql & ", balact = " & l_balact 
		l_sql = l_sql & ", planro = " & l_planro 
		l_sql = l_sql & ", balvpc = " & l_balvpc 
		l_sql = l_sql & ", balmarca = '" & l_balmarca & "'"		
		l_sql = l_sql & ", balconexion = '" & l_balconexion &"'" 		
  	    l_sql = l_sql & " WHERE balnro = " & l_balnro
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

