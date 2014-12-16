<%  Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<%
on error goto 0
'---------------------------------------------------------------------------------
'Archivo	: nota_adp_03.asp
'Descripción: Grabar datos de notas
'Autor		: Claudia Cecilia Rossi
'Fecha		: 30-08-2003
'Modificado	: 08-11-05 - Leticia A. - Adecuarlo para Autogestion.
'----------------------------------------------------------------------------------
%>
<script>
var xc = screen.availWidth;
var yc = screen.availHeight;
window.moveTo(xc,yc);	
window.resizeTo(20,20);</script>
<%
on error goto 0
'Declaracion de Variables locales  -------------------------------------
 Dim l_tipo
 Dim l_cm
 Dim l_sql
 dim l_rs

 Dim l_tnonro
 Dim l_notanro
 Dim l_notatxt
 Dim l_ternro
 'Dim l_tnoconfidencial
 Dim l_tnonroant
 
 Dim l_notfecalta
 Dim l_nothoraalta_h
 Dim l_nothoraalta_m
 Dim l_nothoraalta
 Dim l_notfecvenc
 Dim l_nothoravenc_h
 Dim l_nothoravenc_m
 Dim l_nothoravenc
 Dim l_notremitente
 Dim l_notmotivo
 
'valores del form de alta/modificacion ----------------------------
 l_ternro = l_ess_ternro
 l_tipo 	 = Request.Form("tipo")
 l_tnonro	 = request.Form("tnonro")
 l_tnonroant = request.Form("tnonroant")
 l_notatxt	 = request.Form("notatxt")
 l_notanro	 = request.Form("notanro")

 l_notfecalta 		= request.Form("notfecalta")
 l_nothoraalta_h 	= request.Form("nothoraalta_h")
 l_nothoraalta_m	= request.Form("nothoraalta_m")
 l_notfecvenc		= request.Form("notfecvenc")
 l_nothoravenc_h	= request.Form("nothoravenc_h")
 l_nothoravenc_m	= request.Form("nothoravenc_m")
 l_notremitente		= request.Form("notremitente")
 l_notmotivo		= request.Form("notmotivo")
 
 function eliminarCaracteres(strWords)
 dim badChars 
 dim newChars
 dim i
   badChars = array("drop", ";", "--", chr(34), "xp_") 
   newChars = strWords
   for i = 0 to uBound(badChars) 
     newChars = replace(newChars, badChars(i), " ") 
   next
   eliminarCaracteres = newChars
end function 'eliminarCaracteres


'armo las horas
 l_nothoraalta		=  l_nothoraalta_h & l_nothoraalta_m
 if l_nothoravenc_h = "" or  l_nothoravenc_m = "" then
	  l_nothoravenc_m = "null"
	else
	  l_nothoravenc		=  l_nothoravenc_h & l_nothoravenc_m
 end if
 
 if l_notfecvenc = "" then 
 	   l_notfecvenc = "null" 
	else 
	   l_notfecvenc = cambiafecha(l_notfecvenc,"YMD",true)
 end if
 
 'acomodar las descripciones extendidas
' if len(l_notatxt) <> 0 then
'	l_notatxt = left(trim(l_notatxt),200)
' end if
 
'Body
 set l_cm = Server.CreateObject("ADODB.Command")
 if l_tipo = "A" then 
		l_sql = "INSERT INTO notas_ter "
		l_sql = l_sql & "(ternro, "
		l_sql = l_sql & " tnonro, "
		l_sql = l_sql & " notatxt, "
		l_sql = l_sql & " notfecalta, "
		l_sql = l_sql & " nothoraalta, "
		l_sql = l_sql & " notfecvenc, "
		l_sql = l_sql & " nothoravenc, "
		l_sql = l_sql & " notremitente, "
		l_sql = l_sql & " notmotivo ) "
		
 		l_sql = l_sql & " values (" 
		l_sql = l_sql & l_ternro & ","
		l_sql = l_sql & l_tnonro & ",'"
		l_sql = l_sql & l_notatxt & "',"
		l_sql = l_sql & cambiafecha(l_notfecalta,"YMD",true) & ",'"
		l_sql = l_sql & l_nothoraalta & "',"
		l_sql = l_sql & l_notfecvenc & ",'"
		l_sql = l_sql & l_nothoravenc & "','"
		l_sql = l_sql & l_notremitente & "','"
		l_sql = l_sql & l_notmotivo & "')"
		
		l_cm.activeconnection = Cn
		l_cm.CommandText = eliminarCaracteres(l_sql)
		l_cm.Execute
		'cmExecute l_cm, l_sql, 0	
		Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
else
		l_sql = "UPDATE notas_ter SET "
		l_sql = l_sql & " notatxt   ='" & l_notatxt & "', "
		l_sql = l_sql & " tnonro    = " & l_tnonro & ", "
		
		l_sql = l_sql & " notfecalta = " & cambiafecha(l_notfecalta,"YMD",true) & ","
		l_sql = l_sql & " nothoraalta = '" & l_nothoraalta & "',"
		l_sql = l_sql & " notfecvenc = " & l_notfecvenc & ","
		l_sql = l_sql & " nothoravenc = '" & l_nothoravenc & "',"
		l_sql = l_sql & " notremitente = '" & l_notremitente & "',"
		l_sql = l_sql & " notmotivo = '" & l_notmotivo & "' "
		
		l_sql = l_sql & " WHERE notanro = " & l_notanro
		l_cm.activeconnection = Cn
		l_cm.CommandText = eliminarCaracteres(l_sql)
		l_cm.Execute
		'cmExecute l_cm, l_sql, 0	
		'Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'nota_adp_01.asp?ternro="&l_ternro& "&tnoconfidencial=" &l_tnoconfidencial &"';window.close();</script>"
		Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
end if
%>
