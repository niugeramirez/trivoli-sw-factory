<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: embarque_con_06.asp
'Descripción: ABM de Embarque
'Autor : Gustavo manfrin
'Fecha: 18/09/2006

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_embnro
Dim l_embcod
Dim l_embact
Dim l_carconnro
Dim l_ordnro
Dim l_connro
Dim l_desnro
Dim l_cornro

Dim texto
Dim pronro
Dim empnro
Dim vennro
Dim cornro


texto 		= ""
l_tipo		= request.QueryString("tipo")
l_embnro    = request.QueryString("embnro")
l_embcod    = request.QueryString("embcod")
l_embact	= request.QueryString("embact")
l_ordnro	= request.QueryString("ordnro")
l_connro	= request.QueryString("connro")
l_desnro	= request.QueryString("desnro")
l_cornro	= request.QueryString("cornro")


'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT embnro"
l_sql = l_sql & " FROM tkt_embarque "
l_sql = l_sql & " WHERE embcod=" & l_embcod 
if l_tipo = "M" then
	l_sql = l_sql & " AND embnro <> " & l_embnro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
   	texto =  "Ya existen estos datos."
else
    if l_embact then
		l_rs.close
	   	l_sql = "SELECT embnro"
		l_sql = l_sql & " FROM tkt_embarque "
		l_sql = l_sql & " WHERE embact = -1"
		if l_tipo = "M" then
			l_sql = l_sql & " AND embnro <> " & l_embnro
		end if
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then
		   	texto =  "Hay otro embarque activo."
		end if
	end if	
end if 

if texto = "" and l_connro <> "0" then
	l_rs.close
	l_sql = "SELECT pronro, empnro, vennro, cornro "
	l_sql = l_sql & " FROM tkt_contrato "
	l_sql = l_sql & " WHERE connro = " & l_connro
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		pronro = l_rs("pronro")		
		empnro = l_rs("empnro")		
		vennro = l_rs("vennro")		
		cornro = l_rs("cornro")				
 		l_rs.close
		l_sql = "SELECT pronro, empnro "
		l_sql = l_sql & " FROM tkt_ordentrabajo "
		l_sql = l_sql & " WHERE ordnro = " & l_ordnro
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then
		   if pronro <> l_rs("pronro") then
   		   		texto =  "El producto es distinto en el Contrato y en la Orden."
		   else
     		   if pronro <> l_rs("pronro") then
    		   		texto =  "La empresa es distinta en el Contrato y en la Orden."
	 	       else
  		   			if trim(cornro) <> trim(l_cornro) then
		   				texto =  "El Corredor no corresponde al Contrato." 
					end if	

			   		if l_desnro <> "" and texto = "" then 
					 	l_rs.close
						l_sql = "SELECT vencornro"
						l_sql  = l_sql  & " FROM tkt_vencor "
						l_sql  = l_sql  & " WHERE tkt_vencor.vencorcod = '" & l_desnro & "'"
						rsOpen l_rs, cn, l_sql, 0
						if not l_rs.eof  then 
 	     		   			if trim(vennro) <> trim(l_rs("vencornro")) then
    			   				texto =  "El Destinatario no corresponde al Contrato."
							end if	
						end if	
				   	end if	   
               end if  
		   end if
		end if   
	end if
end if



l_rs.close
%>

<script>
<% 
 if texto <> "" then
%>
   parent.invalido('<%= texto %>')
<% else%>
   parent.valido();
<% end if%>
</script>

<%
Set l_rs = Nothing
%>

