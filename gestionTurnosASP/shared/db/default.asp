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
 


' Response.write "<script>alert('"&  l_iduser &".');</script>"
' Response.write "<script>alert('"&   l_pass &".');</script>"
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
 <!--#include virtual="/turnos/shared/inc/encrypt.inc"-->
 <!--#include virtual="/turnos/shared/db/conn.inc"-->
<%

 Session("UserName") = "sa" ' l_iduser'
 Session("Password") = "" 'l_pass	'""


' if Cint(l_baseBD) = 2 then
'	' Esto esta cableado para Bahia Blanca, ya que en RHPROX2 no se puede generar el usuario ess
'	Session("UserName") = "sa"
'	Session("Password") = ""
' end if
 Session("base") = l_baseBD
' Session("base_z") = l_baseBD
 Session("Time") = now
 
%>
 <!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<%
  
 l_seguir = true
 l_MsgAdv = false
 l_msgtxt = ""
  
ON ERROR resume next
 if err then
 	' La coneccion a la base de datos da error con el usuario 'ess'
	if l_menu = "html" then
		if l_debug = -1 then
			%><script>parent.document.FormVar.desc.value = "<%= Err.Description %>";</script><%
		else
			%><script>parent.document.FormVar.desc.value = "Acceso no valido";</script><%
		end if
	else
		if l_debug = -1 then
			response.write "&acceso=" & Err.Description & "&"
		else
			response.write "&acceso=Acceso no valido&"
		end if
	end if
	l_seguir = false
 else
 	' La conexion a la base de datos es valida con el usuario 'sa'
	l_seguir = true
 end if
 
 if l_seguir then
 	' Restauro los valores de user y pass a los del usuario, en el caso que no utilice NT
	'if l_seg_NT = 0 then
	'	Session("UserName") = l_iduser
	'	Session("Password") = l_pass
	'end if
	
	' Ingreso en la base de datos el logueo del usuario
	'ingresarlogueo l_iduser
	
	Session("loguinUser") = l_iduser
	
	if l_seg_NT = 0 and CInt(l_cambiarpass) = -1 then
		if l_menu = "html" then
			%><script>parent.location = "../../lanzador/lanzador2.asp?tipo=pass";</script><%
		else
			response.write "&acceso=Valido&"
			response.write "&cambiapass=-1&"
			response.write "&msgtxt=" & l_msgtxt & "&"
		end if
	else
	 	' Restauro los valores de user y pass a los del usuario, en el caso que no utilice NT
		if l_seg_NT = 0 then
			Session("UserName") = l_iduser
			Session("Password") = l_pass
		end if
		
		' Ingreso en la base de datos el logueo del usuario
		ingresarlogueo l_iduser
		
		if l_menu = "html" then
			%><script>//parent.document.location = "../../lanzador/lanzador3.asp";</script><%
			%><script>parent.document.location = "../asp/asistente_00_RH.asp?wiznro=8";</script><%
			%><script>//parent.document.location = "../../config/menu.asp?wiznro=8";</script><%
		else
			if not l_MsgAdv then
				response.write "&acceso=Valido&"
				response.write "&cambiapass=0&"
				response.write "&msgtxt=&"
			end if
		end if
	end if
 else
 	Session.Abandon
 end if

%>

