<%@ Language=VBScript %>
<%
Option Explicit
on error goto 0
dim links
Dim l_version
Dim l_inicio
Dim l_email
Dim l_ayuda
Dim l_empresa
Dim l_conexion
Dim l_seguridad
Dim l_debug
Dim l_menu
Dim l_ventana
Dim l_tipo
Dim l_cant
	l_cant = 0
Dim l_seguridad_integrada
Dim l_cont
Dim l_combo(50)
Dim l_combo_cadena
Dim l_winuser
Dim l_domwinuser

l_menu = request("menu")
l_ventana = request("ventana")
l_tipo = request("tipo")

l_version	= "X2"
l_inicio	= "http://www.heidt.com.ar"
l_email		= "gmanfrin@manera.com.ar"
l_ayuda		= "about:blank"	' Si  no hay ayuda abro una ventana en blanco
l_empresa	= "Desarrollo"	'Oleaginosa Moreno'si empresa <> desa, muestra el label de la empresa, sino muestra el combo con las bases de datos
l_conexion	= 2			' numero para la base en default.asp, si empresa <> "desa"
l_seguridad = 0				' (-1) = seguridad por nt cuando l_empresa <> "desa" y no se muestra el combo
l_debug		= 0			' si la variable debug es verdadera (-1) muestra los errores que encuentre"
l_winuser 	= Request.ServerVariables("LOGON_USER")
			'seguridad: si es -1 indique que se va a utilizar la seguridad integrada con NT
			'	o si es 0 que se va a utilizar la seguridad de la bd.
'combo(n)  = "descripcion, base, seguridad, seleccionada"
'l_combo(0) = "RhProX2,2,0,-1"
'l_combo(1) = "Citrusvil,10,0,0"
'l_combo(2) = "PromofilmNew,8,0,0"
'l_combo(3) = "Uruguay,11,-1,0"
'l_combo(4) = "TTI,12,0,0"
'l_combo(5) = "Flecha Bus,13,0,0"
'l_combo(6) = "Base demo,14,0,0"
'l_combo(7) = "New Prod,15,0,0"


'-------------Fin confguraci�n-------------'
if l_menu <> "html" then
	'for l_cont = 0 to Ubound(l_combo)
	'	if l_combo(l_cont) <> "" then
	'		l_combo_cadena = l_combo_cadena & l_combo(l_cont) & ";"
	'	end if
	'next
	links = "&version=" & l_version & "&"
	links = links & "inicio=" & l_inicio & "&"
	links = links & "email=" & l_email & "&"
	links = links & "ayuda=" & l_ayuda & "&"
	links = links & "empresa=" & l_empresa & "&" 
	links = links & "seguridad=" & l_seguridad & "&" 
	links = links & "conexion=" & l_conexion & "&" 
	links = links & "debug=" & l_debug & "&" 
	links = links & "winuser=" & l_winuser & "&" 
	'links = links & "combo=" & l_combo_cadena & "&" 
	'links = links & "&"
	Response.Write links
else %>
	<script>
	//parent.FormVar.version.value = '<%= l_version %>';
	parent.document.FormVar.empresa.value = '<%= l_empresa %>';
	parent.document.FormVar.seguridad.value = <%= l_seguridad %>;
	parent.document.FormVar.conexion.value = <%= l_conexion %>;
	parent.document.FormVar.inicio.value = '<%= l_inicio %>';
	parent.document.FormVar.email2.href = 'mailto:<%= l_email %>';
	parent.document.FormVar.ayuda.value = '<%= l_ayuda %>';
	parent.document.FormVar.debug.value = <%= l_debug %>;
	<%	
	if l_tipo <> "pass" then
		if l_ventana <> "03" then
			dim l_arreglo 
			if l_winuser = "" then
				l_winuser = "An�nimo"
				l_domwinuser = ""
			else
				l_domwinuser = l_winuser
				l_arreglo = split(l_winuser,"\")
				l_winuser = l_arreglo(1)
			end if
			'if l_empresa = "desa" then
				'for l_cont = 0 to Ubound(l_combo)
				'	if l_combo(l_cont) <> "" then 
				'		l_arreglo = split(l_combo(l_cont),",") %>
				//			var opt = document.createElement("OPTION");
				//			opt.text	= "<%'= l_arreglo(0) %>";
				//			opt.value	=  <%'= CInt(l_cont) %>;
				//			opt.seg		= '<%'= l_arreglo(2) %>';
				//			opt.bases	= '<%'= l_arreglo(1) %>';
				//			opt.user	= '<%' if CInt(l_arreglo(2)) = -1 then response.write l_winuser end if%>';
							<%' If l_arreglo(3) then %>
				//			opt.selected= true;
				//			parent.FormVar.seg_nt.value = <%'= l_arreglo(2) %>
				//			parent.FormVar.seguridad.value = <%'= l_arreglo(2) %>
							<%' End If %>
				//			parent.FormVar.basex.add(opt);
				<%'	l_cant = l_cant +1
				'	end if
				'next 
				'if l_cant > 0 then %>
				//	parent.FormVar.basex.remove(0);
				//	parent.document.FormVar.base.value = parent.document.FormVar.basex[parent.document.FormVar.basex.selectedIndex].bases;
				//	if (parent.document.FormVar.basex[parent.document.FormVar.basex.selectedIndex].seg == -1){
				//		parent.document.FormVar.usr2.value = parent.document.FormVar.basex[parent.document.FormVar.basex.selectedIndex].user;
				//		parent.document.FormVar.usr2.disabled = true;
				//		parent.document.FormVar.pass.disabled = true;
				//	}
				<%' End If %>
			<%' Else ' empresa <> "desa" cambio el combo x un texto%>
				/*parent.combobox.style.position ="absolute";
				parent.combobox.style.top="186px";
				parent.combobox.style.left="130px";*/
				//parent.combobox.innerHTML = "<input name=basex readonly seg=<%= l_seguridad %> bases=<%= l_conexion %> user=<%= l_winuser %> style='width:140px;border:none;text-align:left;font-family:Arial;font-size:17px;color:Blue;background-color:transparent;' tabindex=-1 type=Text value='<%= l_empresa %>'>";
				//parent.tere.innerHTML = "<td colspan=2 align=center><input name=basex readonly seg=<%= l_seguridad %> bases=<%= l_conexion %> user=<%= l_winuser %> style='width:140px;border:none;text-align:left;font-family:Arial;font-size:17px;color:Blue;background-color:transparent;' tabindex=-1 type=Text value='<%= l_empresa %>'></td>";
				
				// Raul 28/09/2007 de aca para abajo se comentario
				//parent.document.FormVar.basex.seg = <%'= l_seguridad %>;
				//parent.document.FormVar.basex.bases = <%'= l_conexion %>;
				//parent.document.FormVar.basex.user = '<%'= l_winuser %>';
				//parent.document.FormVar.basex.value = '<%'= l_empresa %>';
				//parent.document.all.texto3.style.visibility ="hidden";
				//parent.document.FormVar.base.value = <%'= l_conexion %>;
				<% 'If l_seguridad = -1 then %>
					//parent.document.FormVar.usr2.value = '<%'= l_winuser %>';
					//parent.document.FormVar.usr2.disabled = true;
					//parent.document.FormVar.pass.disabled = true;
				<% 'End If %>
			<%' End If ' fin empresa%>
			//parent.usertmp = '<%= l_winuser %>';
			//parent.loguser = '';
			//parent.passuser = '';
			//parent.domuser = '<%= replace(l_domwinuser,"\","#@#") %>';
		<%	end if %>
	<% End If %>
	</script>
<% end if %> 