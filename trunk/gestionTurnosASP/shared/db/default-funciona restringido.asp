<%
' Opciones de los parametros que devuelve para el caso de flash
' 	ACCESO		|	 CAMBIAPASS		|	  MSGTXT		|	DESCRIPCION
'---------------|-------------------|-------------------|---------------------------
'	Valido		|		-1			|	Indistinto		|	Cambia la contraseña
'	Valido		|		0			|		''			|	Acceso normal al Sistema (Usuario/contraseña validos)
'	Valido		|		0			|	Un mensaje		|	Muestra un mensaje de Advertencia y permite el ingreso al sistema
'	No Valido	|		0			|	Un mensaje		|	Muestra un mensaje de Error y se queda en el loguin (NO permite el ingreso al sistema)
'
'on error goto 0
 Const c_strEncryptionKey = "56238"

 
 Dim l_pass_expira_dias
 Dim l_pass_camb_dias
 Dim l_pass_int_fallidos
 Dim l_pass_dias_log
 Dim l_usrpasscambiar
 Dim l_seguir
 Dim l_cambiarpass
 Dim l_msgtxt
 Dim l_aux
 Dim l_hlogfecini
 Dim l_hpassfecini
 Dim l_MsgAdv
 Dim l_pol_nro
 
 Dim l_iduser
 Dim l_pass
 Dim l_seg_NT
 Dim l_baseBD
 Dim l_menu
 Dim l_debug
 
 l_iduser	= lcase(Request.Form("usr"))
 l_pass	 	= Request.Form("pass")
 l_seg_NT	= CInt(Request.Form("seg_NT"))
 l_baseBD 	= trim(Request.Form("base"))
 l_menu 	= trim(Request.Form("menu"))
 l_debug 	= CInt(Request.Form("debug"))
 


'Response.write "<script>alert('"&  l_iduser &".');</script>"
'Response.write "<script>alert('"&   l_pass &".');</script>"
'Response.write "<script>alert('"&   l_seg_NT &".');</script>"
'Response.write "<script>alert('"&   l_baseBD &".');</script>"
'Response.write "<script>alert('"&   l_menu &".');</script>"
'Response.write "<script>alert('"&   l_debug &".');</script>"
 
 if l_seg_NT = -1 then
 	l_iduser = replace(l_iduser, "#@#", "\")
 end if
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que imprime mensaje de error.
'------------------------------------------------------------------------------------------------- 
Sub MostrarError (texto)
	if l_menu = "html" then
		%><script>alert('<%= texto %>');</script><%
	else
		response.write "&acceso=No Valido&"
		response.write "&cambiapass=0&"
		response.write "&msgtxt=" & texto & "&"
	end if
end sub
 
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que imprime mensaje de advertencia.
'------------------------------------------------------------------------------------------------- 
Sub MostrarMsgAdv (texto)
	if l_menu = "html" then
		%><script>alert('<%= texto %>');</script><%
	else
		response.write "&acceso=Valido&"
		response.write "&cambiapass=0&"
		response.write "&msgtxt=" & texto & "&"
	end if
end sub
 
 
'------------------------------------------------------------------------------------------------- 
' Bloque principal
'------------------------------------------------------------------------------------------------- 
 
%>
 <!--#include virtual="/serviciolocal/shared/inc/encrypt.inc"-->
 
<%

 Session("UserName") = "sa" ' l_iduser'
 Session("Password") = "" 'l_pass	'""
 
 Response.write "<script>alert('sdaf');</script>"
%>

