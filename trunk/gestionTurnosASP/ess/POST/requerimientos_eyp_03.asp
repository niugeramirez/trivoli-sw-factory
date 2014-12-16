<%  Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<%
on error goto 0

'---------------------------------------------------------------------------------
'Archivo	: requerimientos_eyp_03.asp
'Descripción: Grabar datos de Requerimientos
'Autor		: Raul Chinestra
'Fecha		: 12/09/2006
' Modificado  : 12/09/2006 Raul Chinestra - se agregó Requerimientos de Personal en Autogestión   

'----------------------------------------------------------------------------------
on error goto 0
'Declaracion de Variables locales  -------------------------------------
 Dim l_tipo
 Dim l_cm
 Dim l_sql
 dim l_rs
 
 Dim l_reqpernro	 
 Dim l_reqperdesabr  
 Dim l_reqperdesext  
 Dim l_reqpersolfec  
 Dim l_reqpercanper  
 Dim l_reqperfecalt  
 Dim l_reqperrelpor  
 Dim l_reqperrelfec  
 Dim l_reqperent     
 Dim l_motreqpri     
 Dim l_motprinro     
 Dim l_motreqnro     
 Dim l_puenro        
 Dim l_reqperrep     
 Dim l_reqperremofr  
 Dim l_reqperprires  
 Dim l_reqperpritar  
 Dim l_reqperotrtar  
 Dim l_reqperbenofr  

 Dim l_ternro
 Dim l_empleg
 
'valores del form de alta/modificacion ----------------------------
 l_ternro = l_ess_ternro
 l_empleg = l_ess_empleg
 
 l_tipo 	 = Request.Form("tipo")

 
 l_reqpernro	 = request.Form("reqpernro")
 l_reqperdesabr  = request.Form("reqperdesabr")
 l_reqperdesext  = request.Form("reqperdesext")
 l_reqpersolfec  = request.Form("reqpersolfec")
 l_reqpercanper  = request.Form("reqpercanper")
 l_reqperfecalt  = request.Form("reqperfecalt")
 l_reqperrelpor  = request.Form("reqperrelpor")
 l_reqperrelfec  = request.Form("reqperrelfec")
 l_reqperent     = request.Form("reqperent")
 l_motreqpri     = request.Form("motreqpri")
 l_motprinro     = request.Form("motprinro")
 l_motreqnro     = request.Form("motreqnro")
 l_puenro        = request.Form("puenro")
 l_reqperrep     = request.Form("reqperrep")
 l_reqperremofr  = request.Form("reqperremofr")
 l_reqperprires  = request.Form("reqperprires")
 l_reqperpritar  = request.Form("reqperpritar")
 l_reqperotrtar  = request.Form("reqperotrtar")
 l_reqperbenofr  = request.Form("reqperbenofr")

 
 if l_reqperremofr = "" then
	l_reqperremofr = 0
 end if 

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

'Body
 set l_cm = Server.CreateObject("ADODB.Command")
 if l_tipo = "A" then   
		l_sql = "INSERT INTO pos_reqpersonal "
		l_sql = l_sql & "( "
		l_sql = l_sql & " reqperdesabr, "
		l_sql = l_sql & " reqperdesext, "
		l_sql = l_sql & " reqpersolpor, "
		l_sql = l_sql & " reqpersolfec, "
		l_sql = l_sql & " reqpercanper, "
		l_sql = l_sql & " reqperfecalt, "
		l_sql = l_sql & " reqperrelpor, "
		l_sql = l_sql & " reqperrelfec, "
		l_sql = l_sql & " reqperent, "
		l_sql = l_sql & " motreqpri, "
		l_sql = l_sql & " motprinro, "
		l_sql = l_sql & " motreqnro, "
		l_sql = l_sql & " puenro, "
		l_sql = l_sql & " reqperrep, "
		l_sql = l_sql & " reqperremofr, "
		l_sql = l_sql & " reqperprires, "
		l_sql = l_sql & " reqperpritar, "
		l_sql = l_sql & " reqperotrtar, "
		l_sql = l_sql & " reqperbenofr) "
 		l_sql = l_sql & " values ('" 
		l_sql = l_sql & l_reqperdesabr & "','"
		l_sql = l_sql & l_reqperdesext & "',"
		l_sql = l_sql & l_empleg & ","
		l_sql = l_sql & cambiafecha(l_reqpersolfec,"YMD",true) & ","
		l_sql = l_sql & l_reqpercanper & ","
		l_sql = l_sql & cambiafecha(l_reqperfecalt,"YMD",true) & ","
		l_sql = l_sql & l_reqperrelpor & ","
		l_sql = l_sql & cambiafecha(l_reqperrelfec,"YMD",true) & ",'"
		l_sql = l_sql & l_reqperent & "',"
		l_sql = l_sql & l_motreqpri & ","
		l_sql = l_sql & l_motprinro & ","
		l_sql = l_sql & l_motreqnro & ","
		l_sql = l_sql & l_puenro & ","
		l_sql = l_sql & l_reqperrep & ","
		l_sql = l_sql & l_reqperremofr & ",'"
		l_sql = l_sql & l_reqperprires & "','"
		l_sql = l_sql & l_reqperpritar & "','"
		l_sql = l_sql & l_reqperotrtar & "','"
		l_sql = l_sql & l_reqperbenofr & "')"
	
		' response.write l_sql
		' response.end	
		
		l_cm.activeconnection = Cn
		l_cm.CommandText = eliminarCaracteres(l_sql)
		l_cm.Execute
		'cmExecute l_cm, l_sql, 0	
		Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
else
		l_sql = "UPDATE pos_reqpersonal SET "
		l_sql = l_sql & " reqperdesabr = '" & l_reqperdesabr & "', "
		l_sql = l_sql & " reqperdesext = '" & l_reqperdesext & "', "
		l_sql = l_sql & " reqpersolfec =  " & cambiafecha(l_reqpersolfec,"YMD",true) & ","
		l_sql = l_sql & " reqpercanper =  " & l_reqpercanper & ","
		l_sql = l_sql & " reqperfecalt =  " & cambiafecha(l_reqperfecalt,"YMD",true) & ","	
		l_sql = l_sql & " reqperrelpor =  " & l_reqperrelpor & ","
		l_sql = l_sql & " reqperrelfec =  " & cambiafecha(l_reqperrelfec,"YMD",true) & ","
		l_sql = l_sql & " reqperent = '" & l_reqperent & "',"
		l_sql = l_sql & " motreqpri = " & l_motreqpri & ","
		l_sql = l_sql & " motprinro = " & l_motprinro & ","
		l_sql = l_sql & " motreqnro = " & l_motreqnro & ","
		l_sql = l_sql & " puenro = " & l_puenro & ","
		l_sql = l_sql & " reqperrep = " & l_reqperrep & ","
		l_sql = l_sql & " reqperremofr = " & l_reqperremofr & ","
		l_sql = l_sql & " reqperprires = '" & l_reqperprires & "',"
		l_sql = l_sql & " reqperpritar = '" & l_reqperpritar & "',"
		l_sql = l_sql & " reqperotrtar = '" & l_reqperotrtar & "',"
		l_sql = l_sql & " reqperbenofr = '" & l_reqperbenofr & "'"
		l_sql = l_sql & " WHERE reqpernro = " & l_reqpernro
		
		'response.write l_sql
		'response.end
		
		l_cm.activeconnection = Cn
		l_cm.CommandText = eliminarCaracteres(l_sql)
		l_cm.Execute
		'cmExecute l_cm, l_sql, 0	
		'Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'nota_adp_01.asp?ternro="&l_ternro& "&tnoconfidencial=" &l_tnoconfidencial &"';window.close();</script>"
		Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
end if
%>
