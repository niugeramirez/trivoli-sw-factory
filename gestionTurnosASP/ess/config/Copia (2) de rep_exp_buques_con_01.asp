<% Option Explicit
if request.querystring("excel") then
	Response.AddHeader "Content-Disposition", "attachment;filename=Historial Camion Vagon.xls" 
	Response.ContentType = "application/vnd.ms-excel"
end if
 %>

<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!-----------------------------------------------------------------------------
Archivo: rep_c_descargas_cyr_01.asp 
Creacion: 02/05/2007
Descripcion: Control descargas HP9000
------------------------------------------------------------------------------->
<% 
on error goto 0

Const l_Max_Lineas_X_Pag = 50
Const l_cantcols = 10

Dim l_indice
Dim CAS(50)
Dim INF(50)

Dim l_rs
Dim l_rs2
Dim l_buqdes
Dim l_canbuq
Dim l_totton

Dim l_sql
dim primero
dim ultimo

Dim l_enHP9000

Dim l_nrolinea
Dim l_nropagina

Dim l_encabezado
Dim l_corte 

dim l_total 

dim l_fecini
dim l_fecfin
dim l_feciniHP
dim l_fecfinHP
dim l_anulado

dim l_movcod
dim l_operacion

dim l_lugar

'Variable usadas para imprimir los Totales
dim l_nroope

' Imprime los Totales


dim l_rep1
dim l_rep2
dim l_rep3
dim l_rep4
dim l_rep5
dim l_rep6 ' pendiente
dim l_rep7
dim l_rep8


dim l_rep12
dim l_rep13
dim l_rep14

dim l_rep19
dim l_rep21

'---------------
' rep2
'---------------


'---------------
' rep3
Dim l_merdes
Dim l_indice_mercaderia
Dim MatMesMer(50,50)
Dim l_Mes
Dim l_TotMesMer
Dim l_TotTotMesMer
'---------------

'---------------
' rep4
Dim l_indmer
Dim l_totfil
Dim TOTCAS(100)
Dim l_totcaston
Dim y
Dim l_expdes
Dim x
Dim l_existe
Dim l_ColMer
Dim l_TotMerExp
Dim l_TotTotMerExp
'---------------

'---------------
' rep5
Dim l_anioini
'---------------

'---------------
' rep7
Dim l_total_toneladas

Dim l_desdes
Dim ArrDesNro(50)
Dim ArrDesDes(50)
Dim MatMerDes(50,50)
Dim l_indice_destino
Dim l_TotMerDes
Dim l_TotTotMerDes
'---------------

'---------------
' rep8
Dim l_TotMes
Dim l_TotMer
Dim l_TotTotMerMes
'---------------

Dim l_rep7111

l_rep1 = false
l_rep2 = false
l_rep3 = false
l_rep4 = false
l_rep5 = false
l_rep6 = false ' Pendiente
l_rep7 = false ' igual al 5 pero por destino 
l_rep8 = true


l_rep12 = false
l_rep13 = false
l_rep14 = false


l_rep19 = false
l_rep21 = false


Dim l_indice_exportadora


Dim ArrExpNro(50)
Dim ArrExpDes(50)
			
Dim ArrMerNro(50)
Dim ArrMerDes(50)

Dim MatMerExp(50,50)



Function NombreMes(nro)

select case nro
case 1
	NombreMes = "Ene"
case 2
	NombreMes = "Feb"
case 3
	NombreMes = "Mar"
case 4
	NombreMes = "Abr"
case 5
	NombreMes = "May"
case 6
	NombreMes = "Jun"
case 7
	NombreMes = "Jul"
case 8
	NombreMes = "Ago"
case 9
	NombreMes = "Sep"
case 10
	NombreMes = "Oct"
case 11
	NombreMes = "Nov"
case 12
	NombreMes = "Dic"

end select

end Function


sub totales()
	%>
	<tr>
		<td colspan="<%= l_cantcols %>">&nbsp;</td>
	</tr>
	<tr>
		<td align="right"colspan="4"><b>Total de movimientos faltantes en HP9000 :</b></td>
		<td align="center"><b><%= l_Total %></b></td>
		<td align="center" colspan="<%= l_cantcols - 5 %>">&nbsp;</td>					
	</tr>
	<tr>
		<td colspan="<%= l_cantcols %>">&nbsp;</td>
	</tr>		
	<%
end sub 'totales

sub encabezado_expbuq(titulo)
%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="<%= l_cantcols %>">
				<table>
					<tr>
				       	<td nowrap>&nbsp;&nbsp;&nbsp; <%'= Empresa %>
							&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp; <%= l_lugar %>
							&nbsp;
						</td>				

						<td align="center" width="100%">
							<%= l_fecini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>

				       	<td align="right" nowrap > 
							<%= date%>&nbsp;-&nbsp;<%= time%>
						</td>				

						
					</tr>
				</table>
			</td>				
		</tr>
		<tr>
			<td  colspan="<%= l_cantcols %>">&nbsp;</td>
		</tr>		

		<tr>
	    <tr>
	        <th align="center" width="10%">Buque</th>
	        <th align="center" width="10%">Comenzó</th>
			<th align="center" width="10%">Terminó</th>
			<th align="center" width="10%">Toneladas</th>		
	        <th align="center" width="10%">Mercadería</th>
			<th align="center" width="10%">Sitio</th>
			<th align="center" width="10%">Agencia</th>		
			<th align="center" width="10%">Destino</th>																
	    </tr>		
<%
end sub 'encabezado

sub encabezado_impbuq(titulo)
%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="<%= l_cantcols %>">
				<table>
					<tr>
				       	<td nowrap>&nbsp;&nbsp;&nbsp; <%'= Empresa %>
							&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp; <%= l_lugar %>
							&nbsp;
						</td>				

						<td align="center" width="100%">
							<%= l_fecini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>

				       	<td align="right" nowrap > 
							<%= date%>&nbsp;-&nbsp;<%= time%>
						</td>				

						
					</tr>
				</table>
			</td>				
		</tr>
		<tr>
			<td  colspan="<%= l_cantcols %>">&nbsp;</td>
		</tr>		

		<tr>
	    <tr>
	        <th align="center" width="10%">Buque</th>
	        <th align="center" width="10%">Comenzó</th>
			<th align="center" width="10%">Terminó</th>
			<th align="center" width="10%">Toneladas</th>		
	        <th align="center" width="10%">Mercadería</th>
			<th align="center" width="10%">Sitio</th>
			<th align="center" width="10%">Agencia</th>		
			<th align="center" width="10%">Procedencia</th>				
	    </tr>		
<%
end sub 'encabezado

sub encabezado_impmer(titulo)

l_anioini = "01/01/" & year(l_fecfin)

%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="14">
				<table>
					<tr>
				       	<td nowrap>&nbsp;&nbsp;&nbsp; <%'= Empresa %>
							&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				

						<td align="center" width="100%">
							<%= l_anioini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>

				       	<td align="right" nowrap > 
							<%= date%>&nbsp;-&nbsp;<%= time%>
						</td>				

						
					</tr>
				</table>
			</td>				
		</tr>
		<tr>
			<td  colspan="14">&nbsp;</td>
		</tr>		
	    <tr>
	        <th align="center" width="10%">Mercadería</th>		
	        <th align="center" width="10%">ENE</th>
	        <th align="center" width="10%">FEB</th>
			<th align="center" width="10%">MAR</th>
			<th align="center" width="10%">ABR</th>		
	        <th align="center" width="10%">MAY</th>
			<th align="center" width="10%">JUN</th>
			<th align="center" width="10%">JUL</th>		
			<th align="center" width="10%">AGO</th>
			<th align="center" width="10%">SEP</th>
			<th align="center" width="10%">OCT</th>
			<th align="center" width="10%">NOV</th>
			<th align="center" width="10%">DIC</th>												
			<th align="center" width="10%">TON</th>			
	    </tr>		
<%
end sub 'encabezado


sub encabezado_expcas(titulo)
%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="<%= l_cantcols %>">
				<table>
					<tr>
				       	<td nowrap colspan="3">Cámara Portuaria y Marítima de Bahía Blanca
						</td>				
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							Período:&nbsp;&nbsp;<%= l_fecini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>

	    <tr>
	        <th align="center" width="10%">Empresas</th>			
			<%  
			l_sql = " SELECT * "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
			l_sql = l_sql & " inner join buq_exportadora on buq_exportadora.expnro = buq_contenido.expnro "
			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'CAS' "
			l_sql = l_sql & " ORDER BY  buq_exportadora.expdes "
			rsOpen l_rs, cn, l_sql, 0 
			
			'response.write l_sql
			
			if not l_rs.eof then
				l_expdes = ""
			end if
			
			
			l_indice_exportadora = 1
			l_indice_mercaderia = 1
			do while not l_rs.eof
						
				if l_expdes <> l_rs("expdes") then
					ArrExpNro(l_indice_exportadora) = l_rs("expnro")
					ArrExpDes(l_indice_exportadora) = l_rs("expdes")
					l_expdes = l_rs("expdes")
					l_indice_exportadora = l_indice_exportadora + 1
				end if
				
				l_existe = false
				for x = 1 to l_indice_mercaderia - 1
					if l_rs("mernro") = ArrMerNro(x) then
						l_existe = true
						l_ColMer = x
					end if 
				next
				if l_existe = false then
					ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
					ArrMerDes(l_indice_mercaderia) = l_rs("merdes")
					l_ColMer = l_indice_mercaderia
					l_indice_mercaderia = l_indice_mercaderia + 1
				end if 
			
				MatMerExp(l_ColMer , l_indice_exportadora -1) = MatMerExp(l_ColMer , l_indice_exportadora -1) + l_rs("conton")

				l_rs.MoveNext
			loop
			l_rs.Close
			
			
			for x = 1 to l_indice_mercaderia - 1
			%>			  
			   <th align="center" ><%= ArrMerDes(x) %></th>
			<%
			next
			%>			  
			   <th align="center" >Toneladas</th>					
 		    </tr>	
<%
end sub

sub encabezado_expcasanio(titulo)

l_anioini = "01/01/" & year(l_fecfin)

%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="20">
				<table>
					<tr>
				       	<td nowrap colspan="3">Cámara Portuaria y Marítima de Bahía Blanca
						</td>				
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							Período:&nbsp;&nbsp;<%= l_anioini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>

	    <tr>
	        <th align="center" width="10%">Empresas</th>			
			<%  
			l_sql = " SELECT * "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
			l_sql = l_sql & " inner join buq_exportadora on buq_exportadora.expnro = buq_contenido.expnro "
			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'CAS' "
			l_sql = l_sql & " ORDER BY  buq_exportadora.expdes "
			rsOpen l_rs, cn, l_sql, 0 
			
			'response.write l_sql
			
			if not l_rs.eof then
				l_expdes = ""
			end if
			
			
			l_indice_exportadora = 1
			l_indice_mercaderia = 1
			do while not l_rs.eof
						
				if l_expdes <> l_rs("expdes") then
					ArrExpNro(l_indice_exportadora) = l_rs("expnro")
					ArrExpDes(l_indice_exportadora) = l_rs("expdes")
					l_expdes = l_rs("expdes")
					l_indice_exportadora = l_indice_exportadora + 1
				end if
				
				l_existe = false
				for x = 1 to l_indice_mercaderia - 1
					if l_rs("mernro") = ArrMerNro(x) then
						l_existe = true
						l_ColMer = x
					end if 
				next
				if l_existe = false then
					ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
					ArrMerDes(l_indice_mercaderia) = l_rs("merdes")
					l_ColMer = l_indice_mercaderia
					l_indice_mercaderia = l_indice_mercaderia + 1
				end if 
			
				MatMerExp(l_ColMer , l_indice_exportadora -1) = MatMerExp(l_ColMer , l_indice_exportadora -1) + l_rs("conton")

				l_rs.MoveNext
			loop
			l_rs.Close
			
			
			for x = 1 to l_indice_mercaderia - 1
			%>			  
			   <th align="center" ><%= ArrMerDes(x) %></th>
			<%
			next
			%>			  
			   <th align="center" >Toneladas</th>					
 		    </tr>	
<%
end sub


sub encabezado_expcasdes(titulo)

l_anioini = "01/01/" & year(l_fecfin)

%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="20">
				<table>
					<tr>
				       	<td nowrap colspan="3">Cámara Portuaria y Marítima de Bahía Blanca
						</td>				
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							Período:&nbsp;&nbsp;<%= l_anioini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>

	    <tr>
	        <th align="center" width="10%">Empresas</th>			
			<%  
			l_sql = " SELECT * "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
			l_sql = l_sql & " inner join buq_destino on buq_destino.desnro = buq_contenido.desnro "
			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'CAS' "
			l_sql = l_sql & " ORDER BY  buq_destino.desdes "
			rsOpen l_rs, cn, l_sql, 0 
			
			'response.write l_sql
			
			if not l_rs.eof then
				l_desdes = ""
			end if
			
			
			l_indice_destino = 1
			l_indice_mercaderia = 1
			do while not l_rs.eof
						
				if l_desdes <> l_rs("desdes") then
					ArrDesNro(l_indice_destino) = l_rs("desnro")
					ArrDesDes(l_indice_destino) = l_rs("desdes")
					l_desdes = l_rs("desdes")
					l_indice_destino = l_indice_destino + 1
				end if
				
				l_existe = false
				for x = 1 to l_indice_mercaderia - 1
					if l_rs("mernro") = ArrMerNro(x) then
						l_existe = true
						l_ColMer = x
					end if 
				next
				if l_existe = false then
					ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
					ArrMerDes(l_indice_mercaderia) = l_rs("merdes")
					l_ColMer = l_indice_mercaderia
					l_indice_mercaderia = l_indice_mercaderia + 1
				end if 
			
				MatMerDes(l_ColMer , l_indice_destino -1) = MatMerDes(l_ColMer , l_indice_destino -1) + l_rs("conton")

				l_rs.MoveNext
			loop
			l_rs.Close
			
			
			for x = 1 to l_indice_mercaderia - 1
			%>			  
			   <th align="center" ><%= ArrMerDes(x) %></th>
			<%
			next
			%>			  
			   <th align="center" >Toneladas</th>					
 		    </tr>	
<%
end sub


sub encabezado_expinf(titulo)
%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="<%= l_cantcols %>">
				<table>
					<tr>
				       	<td nowrap colspan="3">Cámara Portuaria y Marítima de Bahía Blanca
						</td>				
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							Período:&nbsp;&nbsp;<%= l_fecini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>

	    <tr>
	        <th align="center" width="10%">Empresas</th>			
			<%  
			l_sql = " SELECT distinct(merdes), merord, buq_mercaderia.mernro "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "

			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'INF' "
			l_sql = l_sql & " AND buq_contenido.expnro <> 0 " ' No tengo en cuenta los que no tienen exportadora			
			l_sql = l_sql & "   AND  buq_buque.tipopenro = 3 "
			l_sql = l_sql & " ORDER BY  buq_mercaderia.merord "
			rsOpen l_rs, cn, l_sql, 0 
			l_indice = 0
			do while not l_rs.eof
				INF(l_indice) = l_rs("mernro")
			%>			  
			   <th align="center" ><%= l_rs("merdes") %></th>					
			<%
				 l_indice = l_indice + 1
				 l_rs.movenext
			 loop
			 %>
			 <th align="center" width="10%">Totales</th>
 		    </tr>	
<%

end sub


sub encabezado_expcasaniodestino(titulo)

	l_anioini = "01/01/" & year(l_fecfin)
%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="<%= l_cantcols %>">
				<table>
					<tr>
				       	<td nowrap colspan="3">Cámara Portuaria y Marítima de Bahía Blanca
						</td>				
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							Período:&nbsp;&nbsp;01/01/<%= year(l_fecfin)  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>

	    <tr>
	        <th align="center" width="10%">Destino</th>			
			<%  
			l_sql = " SELECT distinct(merdes), merord, buq_mercaderia.mernro "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "

			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'CAS' "
			l_sql = l_sql & " ORDER BY  buq_mercaderia.merord "
			rsOpen l_rs, cn, l_sql, 0 
			l_indice = 0
			do while not l_rs.eof
				CAS(l_indice) = l_rs("mernro")
			%>			  
			   <th align="center" ><%= l_rs("merdes") %></th>					
			<%
				 l_indice = l_indice + 1
				 l_rs.movenext
			 loop
			 %>
			 <th align="center" width="10%">Toneladas</th>
			 <th align="center" width="10%">%</th>			 
 		    </tr>	
<%
end sub


sub encabezado_cabmarcar(titulo)
%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="<%= l_cantcols %>">
				<table>
					<tr>
				       	<td nowrap>&nbsp;&nbsp;&nbsp; <%'= Empresa %>
							&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp; <%= l_lugar %>
							&nbsp;
						</td>				

						<td align="center" width="100%">
							<%= l_fecini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>

				       	<td align="right" nowrap > 
							<%= date%>&nbsp;-&nbsp;<%= time%>
						</td>				

						
					</tr>
				</table>
			</td>				
		</tr>
		<tr>
			<td  colspan="<%= l_cantcols %>">&nbsp;</td>
		</tr>		

		<tr>
	    <tr>
	        <th align="center" width="10%">Buque</th>
	        <th align="center" width="10%">Comenzó</th>
			<th align="center" width="10%">Terminó</th>
			<th align="center" width="10%">Toneladas</th>		
	        <th align="center" width="10%">Mercadería</th>
			<th align="center" width="10%">Sitio</th>
			<th align="center" width="10%">Agencia</th>		
	    </tr>		
<%
end sub 'encabezado



sub encabezado_detcargasitio(titulo)

%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="20">
				<table>
					<tr>
				       	<td nowrap colspan="3">Cámara Portuaria y Marítima de Bahía Blanca
						</td>				
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							Período:&nbsp;&nbsp;<%= l_fecini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>


<%
end sub





sub encabezado_detatebuqage(titulo)

l_anioini = "01/01/" & year(l_fecfin)

%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="<%= l_cantcols %>">
				<table>
					<tr>
				       	<td nowrap colspan="3">Cámara Portuaria y Marítima de Bahía Blanca
						</td>				
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							Período:&nbsp;&nbsp;<%= l_fecini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>
	    <tr>
	        <td align="center" width="40%">&nbsp;</td>					
	        <th align="center" width="10%">Agencias</th>			
   		    <th align="center" width="10%">Buques Atendidos</th>	
	        <td align="center" width="40%">&nbsp;</td>								 
	    </tr>

<%
end sub


sub encabezado_MovBuqSitMes(titulo)

l_anioini = "01/01/" & year(l_fecfin)

%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="<%= l_cantcols %>">
				<table>
					<tr>
				       	<td nowrap colspan="3">Cámara Portuaria y Marítima de Bahía Blanca
						</td>				
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							<!--
							Período:&nbsp;&nbsp;<%'= l_fecini  %>&nbsp;-&nbsp;<%'= l_fecfin %>
							-->
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>


<%
end sub



sub fin_encabezado
%>
</table>	
<%
end sub 'finencabezado



'Obtengo los parametros
l_fecini 	  = request.querystring("qfecini")
l_fecfin 	  = request.querystring("qfecfin")

l_anioini = "01/01/" & year(l_fecfin)


l_anulado 	  = request.querystring("anulado")
if l_anulado = "" then
   l_anulado = "false"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<% If not request.querystring("excel") then %>
	<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<% End If %>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")


if l_rep1 = true then

encabezado_expbuq("Exportación - Detalle de Buques") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0


l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " left join buq_destino on buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE  buq_buque.tipopenro = 3 "  ' EXPORTACION
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & " ORDER BY buq_buque.buqfechas, buq_buque.buqfecdes " 
rsOpen l_rs, cn, l_sql, 0

'response.write l_sql
'response.end
if not l_rs.eof then
	l_buqdes = ""
end if

l_canbuq = 0
l_totton = 0
do until l_rs.eof
		%>
		<tr>
			<% if l_buqdes <> l_rs("buqdes") then
			   %>
				<td align="left" width="10%"  nowrap><%=l_rs("buqdes")%></td>			
			   <%
			    l_buqdes = l_rs("buqdes")
				l_canbuq = l_canbuq + 1
			   else
			   %>
				<td align="left" width="10%"  nowrap>&nbsp;</td>			
			   <%
  			   end if
			 %>

			<td align="center" width="10%" ><%= l_rs("buqfecdes") %></td>
			<td align="center" width="10%" ><%= l_rs("buqfechas") %></td>			
			<td align="center" width="10%" ><%= l_rs("conton") %></td>
			<td align="center" width="10%" ><%= l_rs("merdes") %></td>			
			<td align="center" width="10%" ><%= l_rs("sitdes") %></td>			
			<td align="center" width="10%" ><%= l_rs("agedes") %></td>			
			<td align="center" width="10%" ><%= l_rs("desdes") %></td>			
	    </tr>
		<%
		l_totton = l_totton + l_rs("conton")
		l_buqdes = l_rs("buqdes")
		
	l_rs.MoveNext
loop
l_rs.Close

%>
<tr>
	<td align="center" width="10%" colspan="2" >Cantidad de Buques</td>			
	<td align="center" width="10%" ><b><%= l_canbuq %></b></td>
	<td align="center" width="10%" colspan="2" >Total Toneladas</td>				
	<td align="center" width="10%" ><b><%= l_totton %></b></td>
	<td align="center" width="10%" >&nbsp;</td>			
	<td align="center" width="10%" >&nbsp;</td>			
</tr>
<%

end if 

'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep2 = true then

encabezado_impbuq("Importación - Detalle de Buques") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT buqdes, buqfecdes, buqfechas, buq_mercaderia.merdes, buq_sitio.sitdes, buq_agencia.agedes, buq_destino.desdes, sum(conton) tons "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " left join buq_destino on buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE  buq_buque.tipopenro = 4 "  ' IMPORTACION
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

l_sql = l_sql & " group by buqdes, buqfecdes, buqfechas, buq_mercaderia.merdes, buq_sitio.sitdes, buq_agencia.agedes, buq_destino.desdes "
l_sql = l_sql & " order by buqfechas, buqfecdes "


rsOpen l_rs, cn, l_sql, 0

'response.write l_sql
'response.end
if not l_rs.eof then
	l_buqdes = ""
end if

l_canbuq = 0
l_totton = 0
do until l_rs.eof
		%>
		<tr>
			<% if l_buqdes <> l_rs("buqdes") then
			   %>
				<td align="left" width="10%"  nowrap><%=l_rs("buqdes")%></td>			
			   <%
			    l_buqdes = l_rs("buqdes")
				l_canbuq = l_canbuq + 1
			   else
			   %>
				<td align="left" width="10%"  nowrap>&nbsp;</td>			
			   <%
  			   end if
			 %>

			<td align="center" width="10%" ><%= l_rs("buqfecdes") %></td>
			<td align="center" width="10%" ><%= l_rs("buqfechas") %></td>			
			<td align="center" width="10%" ><%= l_rs("tons") %></td>
			<td align="center" width="10%" ><%= l_rs("merdes") %></td>			
			<td align="center" width="10%" ><%= l_rs("sitdes") %></td>			
			<td align="center" width="10%" ><%= l_rs("agedes") %></td>			
			<td align="center" width="10%" ><%= l_rs("desdes") %></td>			
	    </tr>
		<%
		l_totton = l_totton + l_rs("tons")
		l_buqdes = l_rs("buqdes")
		
	l_rs.MoveNext
loop
l_rs.Close

%>
<tr>
	<td align="center" width="10%" colspan="2" >Cantidad de Buques</td>			
	<td align="center" width="10%" ><b><%= l_canbuq %></b></td>
	<td align="center" width="10%" colspan="2" >Total Toneladas</td>				
	<td align="center" width="10%" ><b><%= l_totton %></b></td>
	<td align="center" width="10%" >&nbsp;</td>			
	<td align="center" width="10%" >&nbsp;</td>			
</tr>
<%

end if 

'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep3 = true then 

encabezado_impmer("Importación - Detalle de Mercaderías") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "

l_sql = l_sql & " WHERE  buq_buque.tipopenro = 4 "  ' IMPORTACION
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & " ORDER BY buq_mercaderia.merdes " 

rsOpen l_rs, cn, l_sql, 0

'response.write l_sql
'response.end

if not l_rs.eof then
	l_merdes = ""
end if

l_indice_mercaderia = 1

do until l_rs.eof

	if l_merdes <> l_rs("merdes") then
		ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
		ArrMerDes(l_indice_mercaderia) = l_rs("merdes")
		
		MatMesMer(month(l_rs("buqfechas")) , l_indice_mercaderia) = MatMesMer(month(l_rs("buqfechas")) , l_indice_mercaderia) + l_rs("conton")		
		
		l_merdes = l_rs("merdes")
		l_indice_mercaderia = l_indice_mercaderia + 1
	else 
		MatMesMer(month(l_rs("buqfechas")) , l_indice_mercaderia) = MatMesMer(month(l_rs("buqfechas")) , l_indice_mercaderia) + l_rs("conton")						
	end if

	l_rs.MoveNext
loop
l_rs.Close

for i = 1 to l_indice_mercaderia - 1
%>
<tr>
	<td align="center" width="10%"><%= ArrMerDes(i) %></td>			
<%	
	l_TotMesMer = 0
	for l_Mes = 1 to 12
	%>
	<td align="right"  width="7%"><%= MatMesMer(l_Mes, i ) %></td>			
	<%
		l_TotMesMer = l_TotMesMer + MatMesMer(l_Mes, i ) 
	next
%>
	<td align="right" width="10%"><%= l_TotMesMer %></td>			
</tr>	
<%
next

%>
<tr>	
	<td align="center" width="10%">Total</td>			
<%

' Totales
l_TotTotMesMer = 0
for l_Mes = 1 to 12
	l_TotMesMer = 0
	for i = 1 to l_indice_mercaderia - 1
		l_TotMesMer = l_TotMesMer + MatMesMer(l_Mes, i ) 
	next
%>	
	<td align="right"  width="7%"><% if l_TotMesMer = 0 then response.write "" else response.write l_TotMesMer end if  %></td>			
<%	
	l_TotTotMesMer = l_TotTotMesMer + l_TotMesMer
next
%>	
	<td  align="right" width="10%"><%= l_TotTotMesMer %></td>			
</tr>	
<%
response.write "</table><p style='page-break-before:always'></p>"
end if



'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep4 = true then 

encabezado_expcas("Exportación de Cereales, Aceites y Subproductos") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

for x = 1 to l_indice_exportadora - 1
	%>
	<tr>
		<td nowrap align="center" width="10%" ><%= ArrExpDes(x) %></td>			
	<%
	l_TotMerExp = 0
	for y = 1 to l_indice_mercaderia - 1
	%>
		<td align="right"  width="10%" ><%= MatMerExp(y,x) %></td>			
	<%
		l_TotMerExp = l_TotMerExp + MatMerExp(y,x )
	next
	%>
		<td align="right" width="10%"><%= l_TotMerExp %></td>			
	</tr>	
	<%
next

%>
	<tr>
		<td align="center" width="10%" >Total</td>			
<%

' Totales
l_TotTotMerExp = 0
for i = 1 to l_indice_mercaderia - 1
	l_TotMerExp = 0
	for x = 1 to l_indice_exportadora - 1
		l_TotMerExp = l_TotMerExp + MatMerExp(i,x)
	next
%>	
	<td align="right"  width="7%"><% if l_TotMerExp = 0 then response.write "" else response.write l_TotMerExp end if  %></td>			
<%	
	l_TotTotMerExp = l_TotTotMerExp + l_TotMerExp
next
%>	
	<td  align="right" width="10%"><%= l_TotTotMerExp %></td>			
</tr>	
<%

response.write "</table><p style='page-break-before:always'></p>"
end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep5 = true then 

encabezado_expcasanio("Exportación de Cereales, Aceites y Subproductos") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

for x = 1 to l_indice_exportadora - 1
	%>
	<tr>
		<td nowrap align="center" width="10%" ><%= ArrExpDes(x) %></td>			
	<%
	l_TotMerExp = 0
	for y = 1 to l_indice_mercaderia - 1
	%>
		<td align="right"  width="10%" ><%= MatMerExp(y,x) %></td>			
	<%
		l_TotMerExp = l_TotMerExp + MatMerExp(y,x )
	next
	%>
		<td align="right" width="10%"><%= l_TotMerExp %></td>			
	</tr>	
	<%
next

%>
	<tr>
		<td align="center" width="10%" >Total</td>			
<%

' Totales
l_TotTotMerExp = 0
for i = 1 to l_indice_mercaderia - 1
	l_TotMerExp = 0
	for x = 1 to l_indice_exportadora - 1
		l_TotMerExp = l_TotMerExp + MatMerExp(i,x)
	next
%>	
	<td align="right"  width="7%"><% if l_TotMerExp = 0 then response.write "" else response.write l_TotMerExp end if  %></td>			
<%	
	l_TotTotMerExp = l_TotTotMerExp + l_TotMerExp
next
%>	
	<td  align="right" width="10%"><%= l_TotTotMerExp %></td>			
</tr>	
<%

response.write "</table><p style='page-break-before:always'></p>"
end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep7 = true then 

encabezado_expcasdes("Total Exportado por Destino - Cereales, Aceites y Subproductos") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

for x = 1 to l_indice_destino - 1
	%>
	<tr>
		<td nowrap align="center" width="10%" ><%= ArrDesDes(x) %></td>			
	<%
	l_TotMerDes = 0
	for y = 1 to l_indice_mercaderia - 1
	%>
		<td align="right"  width="10%" ><%= MatMerDes(y,x) %></td>			
	<%
		l_TotMerDes = l_TotMerDes + MatMerDes(y,x)
	next
	%>
		<td align="right" width="10%"><%= l_TotMerDes %></td>			
	</tr>	
	<%
next
%>
	<tr>
		<td align="center" width="10%" >Total</td>			
<%

'response.end

' Totales
l_TotTotMerDes = 0
for i = 1 to l_indice_mercaderia - 1
	l_TotMerDes = 0
	for x = 1 to l_indice_destino - 1
		l_TotMerDes = l_TotMerDes + MatMerDes(i,x)
	next
%>	
	<td align="right"  width="7%"><% if l_TotMerDes = 0 then response.write "" else response.write l_TotMerDes end if  %></td>			
<%	
	l_TotTotMerDes = l_TotTotMerDes + l_TotMerDes
next
%>	
	<td  align="right" width="10%"><%= l_TotTotMerDes %></td>			
</tr>	
<%
response.write "</table><p style='page-break-before:always'></p>"

end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep8 = true then 

encabezado_detcargasitio("Detalle de Cargas por Sitio") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

Dim ArrMerMes(130, 12)

l_sql = "  SELECT distinct(sitdes) ,buq_contenido.sitnro "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
rsOpen l_rs, cn, l_sql, 0
do while not l_rs.eof

	' Inicializo el Arreglo
	for i = 1 to 100
		for j = 1 to 12
			ArrMerMes(i,j) = 0
		next 
	next

	'response.write l_rs(0) & " - "
	
	l_sql = " SELECT  * "
	l_sql = l_sql & " FROM buq_buque "
	l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
	l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
	l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
	l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
	l_sql = l_sql & " AND buq_contenido.sitnro = " & l_rs(1)
	l_sql = l_sql & " Order by buq_mercaderia.mernro "
	
	l_indice_mercaderia = 1
	l_merdes = ""
	rsOpen l_rs2, cn, l_sql, 0
	do while not l_rs2.eof
		if l_merdes <> l_rs2("merdes") then
			ArrMerNro(l_indice_mercaderia) = l_rs2("mernro")
			ArrMerDes(l_indice_mercaderia) = l_rs2("merdes")
			l_merdes = l_rs2("merdes")
			l_indice_mercaderia =  	l_indice_mercaderia + 1
		end if
		ArrMerMes(l_indice_mercaderia - 1, month(l_rs2("buqfechas"))  ) = ArrMerMes(l_indice_mercaderia - 1 , month(l_rs2("buqfechas"))  )  + l_rs2("conton")
		l_rs2.movenext
	loop
	l_rs2.close
	
	
	%>
	<tr>
	  <td align="center" width="100%">							
		  	<table border="0">
			  <tr>
				  <th align="left" colspan="<%= l_indice_mercaderia + 1 %>" ><%= l_rs("sitdes") %></th>													  
		      </tr>		  
			  <tr>
				  <td align="center" width="5%">Mes</td>							
	<%
	for i = 1 to l_indice_mercaderia - 1
	%>
				  <td align="center" width="5%"><%= ArrMerDes(i) %></td>							
	<%
	next
	%>
				  <td align="center" width="5%">Total</td>	
			  </tr>		  							
	<%
	
	
	for j = 1 to month(l_fecfin)
	%>
			  <tr>
				  <td align="center" width="5%"><%= NombreMes(j) %></td>							
	<%
		l_TotMes = 0
		for i = 1 to l_indice_mercaderia - 1
			%>
				  <td align="center" width="5%"><%= ArrMerMes(i,j) %></td>							
			<%
			l_TotMes = l_TotMes + ArrMerMes(i,j)
		next
	%>
				  <td align="center" width="5%"><%= l_TotMes %></td>								
				</tr>
	<%
	next
	%>
				<tr>		
				  <td align="center" width="5%">Total</td>										
	<%
	l_TotTotMerMes = 0
	for i = 1 to l_indice_mercaderia - 1
		l_TotMer = 0
		for j = 1 to month(l_fecfin)
			l_TotMer = l_TotMer + ArrMerMes(i,j)
		next
		l_TotTotMerMes = l_TotTotMerMes + l_TotMer
		%>
				  <td align="center" width="5%"><%= l_TotMer %></td>							
		<%
	next
	%>
				  <td align="center" width="5%"><%= l_TotTotMerMes %></td>							
				</tr>	  
			</table>
		</td>
	</tr>						
	<%
	
	'response.end
	
	l_rs.movenext
loop
l_rs.close

response.write "</table><p style='page-break-before:always'></p>"

end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep12 = true then 

encabezado_expinf("Exportación Inflamables") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT expdes, buq_mercaderia.mernro, sum(conton) "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " inner join buq_exportadora on buq_exportadora.expnro = buq_contenido.expnro "
l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'INF' "
l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & " AND (buq_buque.tipopenro = 3) "
l_sql = l_sql & " AND buq_contenido.expnro <> 0 " ' No tengo en cuenta los que no tienen exportadora
l_sql = l_sql & " group by expdes, buq_mercaderia.mernro "
rsOpen l_rs, cn, l_sql, 0 

response.write l_sql

do while not l_rs.eof

	'ArrInfExp(l_rs("mernro"), l_rs("expnro")) = ArrInfExp(l_rs("mernro"), l_rs("expnro")) + 
	%>
		<tr>
		<td align="center" width="10%" ><%= l_rs(0) %></td>			
	<%
	l_totfil = 0
	for l_indmer = 0 to l_indice -1
		
		if INF(l_indmer) = l_rs(1) then 
		%>
			<td align="center" width="10%" ><%= l_rs(2) %></td>			
		<%
			l_totfil = l_totfil + l_rs(2)
			TOTCAS(l_indmer) = TOTCAS(l_indmer) + l_rs(2)
		else 
		%>
			<td align="center" width="10%" >&nbsp;</td>			
		<%
		end if
	next
	%>
		<td align="center" width="10%" ><%= l_totfil %></td>			
		</tr>
	<%
	l_rs.movenext
loop
%>
<tr>
	<td align="center"  width="10%" >Total</td>			
<%
l_totcaston = 0
for l_indmer = 0 to l_indice -1
%>
	<td align="center"  width="10%" ><%= TOTCAS(l_indmer) %></td>			
<%
	l_totcaston = l_totcaston + TOTCAS(l_indmer)
	TOTCAS(l_indmer) = 0
next
%>
	<td align="center"  width="10%" ><%= l_totcaston %></td>
</tr>
<%
l_totcaston = 0
response.write "</table><p style='page-break-before:always'></p>"
l_rs.close
end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************


if l_rep13 = true then
encabezado_cabmarcar("Cabotaje Marítimo Nacional - Removido Salidas - Cargas") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0


l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " left join buq_destino on buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE  buq_buque.tipopenro = 1 "  ' CARGAS
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

l_sql = l_sql & " order by buq_buque.buqfechas "

rsOpen l_rs, cn, l_sql, 0

'response.write l_sql
'response.end
if not l_rs.eof then
	l_buqdes = ""
end if

l_canbuq = 0
l_totton = 0
do until l_rs.eof
		%>
		<tr>
			<% if l_buqdes <> l_rs("buqdes") then
			   %>
				<td align="left" width="10%"  nowrap><%=l_rs("buqdes")%></td>			
			   <%
			    l_buqdes = l_rs("buqdes")
				l_canbuq = l_canbuq + 1
			   else
			   %>
				<td align="left" width="10%"  nowrap>&nbsp;</td>			
			   <%
  			   end if
			 %>

			<td align="center" width="10%" ><%= l_rs("buqfecdes") %></td>
			<td align="center" width="10%" ><%= l_rs("buqfechas") %></td>			
			<td align="center" width="10%" ><%= l_rs("conton") %></td>
			<td align="center" width="10%" ><%= l_rs("merdes") %></td>			
			<td align="center" width="10%" ><%= l_rs("sitdes") %></td>			
			<td align="center" width="10%" ><%= l_rs("agedes") %></td>			
	    </tr>
		<%
		l_totton = l_totton + l_rs("conton")
		l_buqdes = l_rs("buqdes")
		
	l_rs.MoveNext
loop
l_rs.Close

%>
<tr>
	<td align="center" width="10%" colspan="2" >Cantidad de Buques</td>			
	<td align="center" width="10%" ><b><%= l_canbuq %></b></td>
	<td align="center" width="10%" colspan="2" >Total Toneladas</td>				
	<td align="center" width="10%" ><b><%= l_totton %></b></td>
	<td align="center" width="10%" >&nbsp;</td>			
</tr>
<%
response.write "</table><p style='page-break-before:always'></p>"
end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************


if l_rep14 = true then
encabezado_cabmarcar("Cabotaje Marítimo Nacional - Removido Entradas - Descargas") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0


l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " left join buq_destino on buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE  buq_buque.tipopenro = 2 "  ' DESCARGAS
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

l_sql = l_sql & " order by buq_buque.buqfechas "

rsOpen l_rs, cn, l_sql, 0

'response.write l_sql
'response.end
if not l_rs.eof then
	l_buqdes = ""
end if

l_canbuq = 0
l_totton = 0
do until l_rs.eof
		%>
		<tr>
			<% if l_buqdes <> l_rs("buqdes") then
			   %>
				<td align="left" width="10%"  nowrap><%=l_rs("buqdes")%></td>			
			   <%
			    l_buqdes = l_rs("buqdes")
				l_canbuq = l_canbuq + 1
			   else
			   %>
				<td align="left" width="10%"  nowrap>&nbsp;</td>			
			   <%
  			   end if
			 %>

			<td align="center" width="10%" ><%= l_rs("buqfecdes") %></td>
			<td align="center" width="10%" ><%= l_rs("buqfechas") %></td>			
			<td align="center" width="10%" ><%= l_rs("conton") %></td>
			<td align="center" width="10%" ><%= l_rs("merdes") %></td>			
			<td align="center" width="10%" ><%= l_rs("sitdes") %></td>			
			<td align="center" width="10%" ><%= l_rs("agedes") %></td>			
	    </tr>
		<%
		l_totton = l_totton + l_rs("conton")
		l_buqdes = l_rs("buqdes")
		
	l_rs.MoveNext
loop
l_rs.Close

%>
<tr>
	<td align="center" width="10%" colspan="2" >Cantidad de Buques</td>			
	<td align="center" width="10%" ><b><%= l_canbuq %></b></td>
	<td align="center" width="10%" colspan="2" >Total Toneladas</td>				
	<td align="center" width="10%" ><b><%= l_totton %></b></td>
	<td align="center" width="10%" >&nbsp;</td>			
</tr>
<%
response.write "</table><p style='page-break-before:always'></p>"
end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep19 = true then 

encabezado_MovBuqSitMes("Movimientos de Buques por Sitio") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")

Dim ArrSitMes(100,100)
Dim i
Dim j
Dim k

for i = 1 to 100
	for j = 1 to 12
		ArrSitMes(i,j) = 0
	next
next

l_sql = " SELECT buq_buque.buqnro, sitnro, buqfechas "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "

l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

'l_sql = l_sql & " AND buq_contenido.sitnro = 1 "
l_sql = l_sql & " group by buq_buque.buqnro, sitnro, buqfechas "
rsOpen l_rs, cn, l_sql, 0

dim valor
Dim l_buquenumero
dim l_sitionumero

do while not l_rs.eof

	ArrSitMes(l_rs("sitnro") , month(l_rs("buqfechas"))  ) = ArrSitMes(l_rs("sitnro") , month(l_rs("buqfechas"))  ) + 1
	
	l_rs.movenext
loop
l_rs.close


l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_sitio "
l_sql = l_sql & " order by buq_sitio.sitnro "

rsOpen l_rs, cn, l_sql, 0
%>
	<tr>
        <th align="center" width="5%">Mes</th>
<%

Dim ArrNomSit(50)
Dim ArrTotSit(50)

i = 0
do while not l_rs.eof
%>
    <th align="center" width="5%"><%= l_rs("sitdes") %></th>
<%
	i = i + 1
	ArrNomSit(i) = l_rs("sitdes")
	l_rs.movenext
loop
l_rs.close
%>
	</tr>
<%

for j = 1 to month(l_fecfin)
%>
  <tr>
	  <td align="center" width="5%"><%= NombreMes(j) %></td>							
<%

	for k = 1 to i
		%>
		  <td align="center" width="5%"><%= ArrSitMes(k,j) %></td>							
		<%
		ArrTotSit(k) = ArrTotSit(k) + ArrSitMes(k,j)
	next
%>
	</tr>
<%
next


%>
  <tr>
	  <td align="center" width="5%"><b>Tot</b></td>							  
<%
for k = 1 to i
%>
	<td align="center" width="5%"><%= ArrTotSit(k) %></td>							
<%
next
%>
  </tr>
<%


encabezado_MovBuqSitMes("Clase y Cantidad de Buques") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")

Dim ArrTipBuqMes(100,100)

for i = 1 to 100
	for j = 1 to 12
		ArrTipBuqMes(i,j) = 0
	next
next

l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_buque "

l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

'l_sql = l_sql & " AND buq_buque.tipbuqnro = 7 "

rsOpen l_rs, cn, l_sql, 0

'response.write l_sql

do while not l_rs.eof

	ArrTipBuqMes(l_rs("tipbuqnro") , month(l_rs("buqfechas"))  ) = ArrTipBuqMes(l_rs("tipbuqnro") , month(l_rs("buqfechas"))  ) + 1
	
	'response.write l_rs("tipbuqnro") & " ---" & month(l_rs("buqfechas")) & "<br>"
	
	'response.write ArrTipBuqMes(2,1) & "<br>"
	
	l_rs.movenext
loop
l_rs.close

'response.write ArrTipBuqMes(2,1)


l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_tipobuque "
l_sql = l_sql & " order by buq_tipobuque.tipbuqnro "
rsOpen l_rs, cn, l_sql, 0
%>
	<tr>
        <th align="center" width="5%">Mes</th>
<%

Dim ArrNomTipBuq(50)
Dim ArrTotTipBuq(50)
Dim ArrTotTipBuqMes(50)

i = 0
do while not l_rs.eof
%>
    <th align="center" width="5%"><%= l_rs("tipbuqdes") %></th>
<%
	i = i + 1
	ArrNomTipBuq(i) = l_rs("tipbuqdes")
	l_rs.movenext
loop
l_rs.close
%>
        <th align="center" width="5%">Totales</th>
	</tr>
<%

for j = 1 to month(l_fecfin)
%>
  <tr>
	  <td align="center" width="5%"><%= NombreMes(j) %></td>							
<%

	for k = 1 to i
		%>
		  <td align="center" width="5%"><%= ArrTipBuqMes(k,j) %></td>							
		<%
		ArrTotTipBuq(k) = ArrTotTipBuq(k) + ArrTipBuqMes(k,j)
		ArrTotTipBuqMes(j) = ArrTotTipBuqMes(j) + ArrTipBuqMes(k,j)
	next
		%>
		  <td align="center" width="5%"><%= ArrTotTipBuqMes(j) %></td>							
		<%
%>
	</tr>
<%
next


%>
  <tr>
	  <td align="center" width="5%"><b>Tot</b></td>							  
<%
Dim l_tottot
l_tottot = 0

for k = 1 to i
%>
	<td align="center" width="5%"><%= ArrTotTipBuq(k) %></td>							
<%
	l_tottot = l_tottot + ArrTotTipBuq(k) 
next
%>
	<td align="center" width="5%"><%= l_tottot %></td>							
  </tr>
<%


response.write "</table><p style='page-break-before:always'></p>"
end if




'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep21 = true then 

encabezado_detatebuqage("Detalle de Atención Buques por Agencia") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT distinct(buq_agencia.agedes), count(*) "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

l_sql = l_sql & " group by buq_agencia.agedes "


rsOpen l_rs, cn, l_sql, 0
do while not l_rs.eof
%>
		<tr>
        <td align="center" width="40%">&nbsp;</td>							
		<td align="center" width="10%" ><%= l_rs(0) %></td>			
		<td align="center" width="10%" ><%= l_rs(1) %></td>					
        <td align="center" width="40%">&nbsp;</td>							
		</tr>		
<%
	l_rs.movenext
loop

response.write "</table><p style='page-break-before:always'></p>"
l_rs.close
end if



'if l_total = 0 then
'   if l_encabezado then 
'	  if l_corte then
'          response.write "</table><p style='page-break-before:always'></p>"
'		  l_nrolinea = 1
'  	  end if 		
'	  encabezado "Control de Movimientos en HP9000" 
'  	  l_nrolinea = l_nrolinea + 3
'	end if 'encabezado
'    l_nrolinea = l_nrolinea + 3
'	response.write ("<tr><td colspan=" & l_cantcols & "><b>No Existen Movimientos para el filtro seleccionado.</b></td></tr>")
'else
'    totales	
'end if
'fin_encabezado

'l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</body>
</html>

