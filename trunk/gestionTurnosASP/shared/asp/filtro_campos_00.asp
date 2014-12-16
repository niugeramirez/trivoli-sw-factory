<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!-------------------------------------------------------------------------------------------
Archivo		: filtro_campos_00.asp
Descripción : Permite realizar un filtrado 
Autor		: Lic. Fernando Favre
Fecha		: 01/2004
Modificado	: 
---------------------------------------------------------------------------------------------
-->

<% 
 Dim l_sql
 Dim l_rs
 
 Dim l_cantidad
 Dim l_actual
 Dim l_etiqueta
 Dim l_campos
 Dim l_tipos
 Dim l_et
 Dim l_camp
 Dim l_tip
 Dim l_liste(20)
 Dim l_listc(20)
 Dim l_listt(20)
 Dim l_i
 Dim l_orden
 Dim l_funct
 Dim l_cclave
 
 l_sql    	= request.Form("sql")
 l_etiqueta = request.Form("etiqueta")
 l_campos 	= request.Form("campos")
 l_tipos  	= request.Form("tipos")
 l_orden 	= request.Form("orden")
 l_funct 	= request.Form("funct")
 l_cclave	= request.Form("campoclave")
 
 l_cantidad = 0
 l_et = l_etiqueta
 do while len(l_et) > 0
 	if inStr(l_et,";") <> 0 then
    	l_actual   = left(l_et, inStr(l_et,";") - 1)
	    l_et = mid (l_et, inStr(l_et,";") + 1)
  	else
    	l_actual = l_et
		l_et = ""
	end if
  	l_cantidad = l_cantidad + 1
	l_liste(l_cantidad) = l_actual
 loop
 
 l_camp = l_campos
 l_cantidad = 0
 do while len(l_camp) > 0
 	if inStr(l_camp,";") <> 0 then
    	l_actual = left(l_camp, inStr(l_camp,";") - 1)
	    l_camp = mid (l_camp, inStr(l_camp,";") + 1)
	else
    	l_Actual = l_camp
		l_camp = ""
  	end if
	l_cantidad = l_cantidad + 1
	l_listc(l_cantidad) = l_actual
 loop
 
 l_tip = l_tipos
 l_cantidad = 0
 do while len(l_tip) > 0
 	if inStr(l_tip,";") <> 0 then
    	l_actual = left(l_tip, inStr(l_tip,";") - 1)
		l_tip = mid (l_tip, inStr(l_tip,";") + 1)
	else
    	l_Actual = l_tip
		l_tip = ""
	end if
	l_cantidad = l_cantidad + 1
	l_listt(l_cantidad) = l_actual
 loop
 
%>
<html>	
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Filtro - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function ActFiltro(tipo, campo){
	if (tipo == 'R'){
		document.all.ifrm_datos.filtro.value = '';
		Filtrar();
	}
	document.ifrm2.location.reload ('filtro_campos_02.asp?tipo=' + tipo + "&campo=" + campo);
}

function Filtrar(){
	document.all.ifrm_datos.submit();
}

function Salir(){
<%	if l_funct <> "" then %>
	if (document.ifrm.document.all.cabnro.value != '')
		if (window.opener.<%= l_funct %>)
			window.opener.<%= l_funct %>(document.ifrm.document.all.cabnro.value);
<%  end if %>

	window.close();
}
</script>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="Filtrar();">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr style="border-color :CadetBlue;">    	
   	<td class="barra">Filtro</td>
	<td class="barra" align="right">		  
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
  <td colspan="2">
	<table border="1" cellspacing="0" cellpadding="0" height="30">
	  	<td valign="top" width="35%">
		  <table border="0" cellspacing="0" cellpadding="0">
	  		<tr>
				<th colspan="2" align="center"><b>Campos</b></th>
			</tr>
			<%
			l_i = 1
			do while l_i <= l_cantidad
			%>  
			<tr>
			    <td align="right"><b><%= l_liste(l_i)%></b></td>
				<td><input type="Radio" name="orden" value="<%= l_listt(l_i) & l_listc(l_i)%>" <% If l_i = 1 then%>checked<% End If %> onclick="ActFiltro('<%= l_listt(l_i) %>', '<%= l_listc(l_i) %>');"></td>
			</tr>
			<%
			  l_i = l_i + 1
			loop
			%>  
			<tr>
			    <td align="right"><b>Restaurar:</b></td>
				<td><input type="Radio" name="orden" value="R" onclick="ActFiltro('R');"></td>
			</tr>
		  </table>
		</td>
		<td valign="top" width="65%" nowrap align="center">
		  <table cellspacing="0" border="0" cellpadding="0" width="100%">
		  	<tr>
    			<th>Filtro</th>
  			</tr>
			<tr> 
				<td align="Center">	    
					<iframe name="ifrm2" src="filtro_campos_02.asp?tipo=<%= l_listt(1) %>&campo=<%= l_listc(1) %>" scrolling="No" frameborder="0"></iframe> 
				</td>
			</tr>
		  </table>	     
		</td>
	</table>	     
  </td>
</tr>
<tr valign="top">
  <td colspan="2" height="100%">
	<iframe name="ifrm" src="#" width="100%"  height="100%" scrolling="Yes"></iframe> 
  </td>
</tr>

<tr>
  <td colspan="2" align="right" class="th2" valign="middle">
	<% call MostrarBoton("sidebtnABM", "Javascript:Salir();", "Aceptar")%>
	<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
  </td>
</tr>

</table>	     

<form id="ifrm_datos" action="filtro_campos_01.asp" method="post" target="ifrm">
  <input type="Hidden" name="tipos"      value="<%=l_tipos%>">
  <input type="Hidden" name="etiqueta"   value="<%=l_etiqueta%>">
  <input type="Hidden" name="campos"     value="<%=l_campos%>">  
  <input type="Hidden" name="campoclave" value="<%=l_cclave%>">
  <input type="Hidden" name="sql"	     value="<%=l_sql%>">
  <input type="Hidden" name="funct"      value="<%=l_funct%>">
  <input type="Hidden" name="filtro"     value="">
  <input type="Hidden" name="orden"	     value="<%=l_orden%>">
</form>
   
<%
 set l_rs = nothing
 Cn.Close
 set Cn = nothing
%>
</body>
</html>
