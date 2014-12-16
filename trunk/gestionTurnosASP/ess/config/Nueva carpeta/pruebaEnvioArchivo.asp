<%@ LANGUAGE = VBScript %>
<%
	Option Explicit

	' Variables de aplicacion
	Dim strServer, strPathDsi, strAp


strAp = "http://cottest.ec.gba.gov.ar/"

response.write strAp + "TransporteBienes/SeguridadCliente/presentarRemitos.do"

	
%>
<HTML>
<HEAD>
	<TITLE>Dirección de Sistemas de Información - Ministerio de Economía de la Provincia de Buenos Aires</TITLE>
	<LINK rel="stylesheet" type="text/css" href="../../includes/aplicaciones.css"></LINK>
</HEAD>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" background="../../Informacion/Templates/ImagenesRentas/I_fondoRG2.jpg">
<DIV>
	<TABLE border="0" cellspacing="0" width="100%">
		<TR> 
            <TD><IMG src="../../Informacion/Templates/ImagenesRentas/EncabezadoChico.jpg" width=780 border=0></TD>             
        </TR>		
		<TR>
			<TD><P align="center">
				<STRONG>Simulación de Transacción de las Empresas de Transporte de Bienes</STRONG></P>
			</TD>
		</TR>
	</TABLE> 
	<br>
	

<FORM action="<%=strAp & "TransporteBienes/SeguridadCliente/presentarRemitos.do"%>" enctype="multipart/form-data" method="post" name="form" onSubmit="">	
	<TABLE border="0" cellSpacing="0" cellPadding="1" width="80%" align="center">

		<TR>
			<TD width="30%" align="right" class="trans">Usuario:&nbsp;</TD>
	   		<TD width="70%" align="left" height="30" class="trans">&nbsp;

											       
	   			<INPUT name="user" maxLength="11" tabindex="" size="20" value="30111111118"> </INPUT>
			</TD>	
		</TR>
		<TR>
			<TD width="30%" align="right" class="trans">Password:&nbsp;</TD>
	   		<TD width="70%" align="left" height="30" class="trans">&nbsp;
	   			<INPUT name="password" maxLength="10" tabindex="" size="20" value="123456"> </INPUT>


			</TD>	
		</TR>
		<TR>
			<TD width="30%" align="right" class="trans">Archivo:&nbsp;</TD>
	   		<TD width="70%" align="left" height="30" class="trans">&nbsp;
	   			<INPUT type="file" name="file" tabIndex=8 style="WIDTH: 
	   			       473px; HEIGHT: 22px" size=36></INPUT>
			</TD>	
		</TR>								
	</TABLE>
	<br>
	<TABLE border="0" cellSpacing="0" cellPadding="1" width="60%" align="center">
		<TR>
			<TD class="trans" align="middle" colspan="2">
				<BR>
				<INPUT type="reset" value="Borrar" tabIndex="9" id=reset1 name=reset1></INPUT>
				<INPUT type="button" value="Volver" tabIndex="11" onclick="javascript:document.location='<%=strPathDsi%>'" id=button1 name=button1></INPUT>										
				<INPUT name="submit" type="submit" value="Enviar" tabIndex="10"></INPUT> 
			</TD>
		</TR>
	</TABLE>	

	
			
		
</DIV>
</FORM>
</body>
</HTML>