<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'Archivo: datos_personal_con_03.asp
'Descripción: Alta de datos de personal
'Autor : Raul Chinestra	
'Fecha: 16/07/2008

Dim l_cm
Dim l_sql
Dim l_tipo

Dim l_ape
Dim l_nom
Dim l_fecnac
Dim l_nac
Dim l_estciv
Dim l_estcivnom
Dim l_pue
Dim l_cal
Dim l_num
Dim l_pis
Dim l_dep
Dim l_loc
Dim l_locdes
Dim l_pro
Dim l_prodes
Dim l_codpos
Dim l_tel
Dim l_cel
Dim l_tipdoc
Dim l_tipdocdes
Dim l_nrodoc
Dim l_cuil
Dim l_apojub
Dim l_apojubdes
Dim l_afjp

Dim l_padape
Dim l_padnom
Dim l_padfecnac
Dim l_padviv
Dim l_padcar
Dim l_padtipdoc
Dim l_padtipdocdes
Dim l_padnrodoc

Dim l_madape
Dim l_madnom
Dim l_madfecnac
Dim l_madviv
Dim l_madcar
Dim l_madtipdoc
Dim l_madtipdocdes
Dim l_madnrodoc

Dim l_conape
Dim l_connom
Dim l_confecnac
Dim l_conviv
Dim l_concar
Dim l_contipdoc
Dim l_contipdocdes
Dim l_connrodoc

Dim l_hi1ape
Dim l_hi1nom
Dim l_hi1fecnac
Dim l_hi1viv
Dim l_hi1car
Dim l_hi1tipdoc
Dim l_hi1tipdocdes
Dim l_hi1nrodoc

Dim l_hi2ape
Dim l_hi2nom
Dim l_hi2fecnac
Dim l_hi2viv
Dim l_hi2car
Dim l_hi2tipdoc
Dim l_hi2tipdocdes
Dim l_hi2nrodoc

Dim l_hi3ape
Dim l_hi3nom
Dim l_hi3fecnac
Dim l_hi3viv
Dim l_hi3car
Dim l_hi3tipdoc
Dim l_hi3tipdocdes
Dim l_hi3nrodoc

Dim l_hi4ape
Dim l_hi4nom
Dim l_hi4fecnac
Dim l_hi4viv
Dim l_hi4car
Dim l_hi4tipdoc
Dim l_hi4tipdocdes
Dim l_hi4nrodoc

Dim l_hi5ape
Dim l_hi5nom
Dim l_hi5fecnac
Dim l_hi5viv
Dim l_hi5car
Dim l_hi5tipdoc
Dim l_hi5tipdocdes
Dim l_hi5nrodoc

Dim l_hi6ape
Dim l_hi6nom
Dim l_hi6fecnac
Dim l_hi6viv
Dim l_hi6car
Dim l_hi6tipdoc
Dim l_hi6tipdocdes
Dim l_hi6nrodoc

Dim l_priins
Dim l_prides
Dim l_prihas
Dim l_pritit

Dim l_secins
Dim l_secdes
Dim l_sechas
Dim l_sectit

Dim l_terins
Dim l_terdes
Dim l_terhas
Dim l_tertit

Dim l_uniins
Dim l_unides
Dim l_unihas
Dim l_unitit

Dim l_posins
Dim l_posdes
Dim l_poshas
Dim l_postit

Dim l_conins
Dim l_condes
Dim l_conhas
Dim l_contit

Dim l_idinom
Dim l_idilee
Dim l_idihab
Dim l_idiesc

Dim l_idi2nom
Dim l_idi2lee
Dim l_idi2hab
Dim l_idi2esc

Dim l_empant1emp
Dim l_empant1pue
Dim l_empant1des
Dim l_empant1has
Dim l_empant1tar

Dim l_empant2emp
Dim l_empant2pue
Dim l_empant2des
Dim l_empant2has
Dim l_empant2tar

Dim l_empant3emp
Dim l_empant3pue
Dim l_empant3des
Dim l_empant3has
Dim l_empant3tar

Dim l_obs

l_tipo 	    = request.Querystring("tipo")
l_ape 	    = request.Form("ape")
l_nom	    = request.Form("nom")
l_fecnac	= request.Form("fecnac")
l_nac		= request.Form("nac")
l_estciv	= request.Form("estciv")
l_pue   	= request.Form("pue")
l_cal   	= request.Form("cal")
l_num   	= request.Form("num")
l_pis   	= request.Form("pis")
l_dep   	= request.Form("dep")
l_loc   	= request.Form("loc")
l_pro   	= request.Form("pro")
l_codpos   	= request.Form("codpos")
l_tel   	= request.Form("tel")
l_cel   	= request.Form("cel")
l_tipdoc  	= request.Form("tipdoc")
l_nrodoc  	= request.Form("nrodoc")
l_cuil  	= request.Form("cuil")
l_apojub  	= request.Form("apojub")
l_afjp  	= request.Form("afjp")

l_padape  	= request.Form("padape")
l_padnom  	= request.Form("padnom")
l_padfecnac	= request.Form("padfecnac")
l_padviv	= request.Form("padviv")
l_padcar	= request.Form("padcar")
l_padtipdoc	= request.Form("padtipdoc")
l_padnrodoc	= request.Form("padnrodoc")

l_madape  	= request.Form("madape")
l_madnom  	= request.Form("madnom")
l_madfecnac	= request.Form("madfecnac")
l_madviv	= request.Form("madviv")
l_madcar	= request.Form("madcar")
l_madtipdoc	= request.Form("madtipdoc")
l_madnrodoc	= request.Form("madnrodoc")

l_conape  	= request.Form("conape")
l_connom  	= request.Form("connom")
l_confecnac	= request.Form("confecnac")
l_conviv	= request.Form("conviv")
l_concar	= request.Form("concar")
l_contipdoc	= request.Form("contipdoc")
l_connrodoc	= request.Form("connrodoc")

l_hi1ape  	= request.Form("hi1ape")
l_hi1nom  	= request.Form("hi1nom")
l_hi1fecnac	= request.Form("hi1fecnac")
l_hi1viv	= request.Form("hi1viv")
l_hi1car	= request.Form("hi1car")
l_hi1tipdoc	= request.Form("hi1tipdoc")
l_hi1nrodoc	= request.Form("hi1nrodoc")

l_hi2ape  	= request.Form("hi2ape")
l_hi2nom  	= request.Form("hi2nom")
l_hi2fecnac	= request.Form("hi2fecnac")
l_hi2viv	= request.Form("hi2viv")
l_hi2car	= request.Form("hi2car")
l_hi2tipdoc	= request.Form("hi2tipdoc")
l_hi2nrodoc	= request.Form("hi2nrodoc")

l_hi3ape  	= request.Form("hi3ape")
l_hi3nom  	= request.Form("hi3nom")
l_hi3fecnac	= request.Form("hi3fecnac")
l_hi3viv	= request.Form("hi3viv")
l_hi3car	= request.Form("hi3car")
l_hi3tipdoc	= request.Form("hi3tipdoc")
l_hi3nrodoc	= request.Form("hi3nrodoc")

l_hi4ape  	= request.Form("hi4ape")
l_hi4nom  	= request.Form("hi4nom")
l_hi4fecnac	= request.Form("hi4fecnac")
l_hi4viv	= request.Form("hi4viv")
l_hi4car	= request.Form("hi4car")
l_hi4tipdoc	= request.Form("hi4tipdoc")
l_hi4nrodoc	= request.Form("hi4nrodoc")

l_hi5ape  	= request.Form("hi5ape")
l_hi5nom  	= request.Form("hi5nom")
l_hi5fecnac	= request.Form("hi5fecnac")
l_hi5viv	= request.Form("hi5viv")
l_hi5car	= request.Form("hi5car")
l_hi5tipdoc	= request.Form("hi5tipdoc")
l_hi5nrodoc	= request.Form("hi5nrodoc")

l_hi6ape  	= request.Form("hi6ape")
l_hi6nom  	= request.Form("hi6nom")
l_hi6fecnac	= request.Form("hi6fecnac")
l_hi6viv	= request.Form("hi6viv")
l_hi6car	= request.Form("hi6car")
l_hi6tipdoc	= request.Form("hi6tipdoc")
l_hi6nrodoc	= request.Form("hi6nrodoc")

l_priins  	= request.Form("priins")
l_prides  	= request.Form("prides")
l_prihas	= request.Form("prihas")
l_pritit	= request.Form("pritit")

l_secins  	= request.Form("secins")
l_secdes  	= request.Form("secdes")
l_sechas	= request.Form("sechas")
l_sectit	= request.Form("sectit")

l_terins  	= request.Form("terins")
l_terdes  	= request.Form("terdes")
l_terhas	= request.Form("terhas")
l_tertit	= request.Form("tertit")

l_uniins  	= request.Form("uniins")
l_unides  	= request.Form("unides")
l_unihas	= request.Form("unihas")
l_unitit	= request.Form("unitit")

l_posins  	= request.Form("posins")
l_posdes  	= request.Form("posdes")
l_poshas	= request.Form("poshas")
l_postit	= request.Form("postit")

l_conins  	= request.Form("conins")
l_condes  	= request.Form("condes")
l_conhas	= request.Form("conhas")
l_contit	= request.Form("contit")

l_idinom  	= request.Form("idinom")
l_idilee  	= request.Form("idilee")
l_idihab	= request.Form("idihab")
l_idiesc	= request.Form("idiesc")

l_idi2nom  	= request.Form("idi2nom")
l_idi2lee  	= request.Form("idi2lee")
l_idi2hab	= request.Form("idi2hab")
l_idi2esc	= request.Form("idi2esc")

l_empant1emp = request.Form("empant1emp")
l_empant1pue = request.Form("empant1pue")
l_empant1des = request.Form("empant1des")
l_empant1has = request.Form("empant1has")
l_empant1tar = request.Form("empant1tar")

l_empant2emp = request.Form("empant2emp")
l_empant2pue = request.Form("empant2pue")
l_empant2des = request.Form("empant2des")
l_empant2has = request.Form("empant2has")
l_empant2tar = request.Form("empant2tar")

l_empant3emp = request.Form("empant3emp")
l_empant3pue = request.Form("empant3pue")
l_empant3des = request.Form("empant3des")
l_empant3has = request.Form("empant3has")
l_empant3tar = request.Form("empant3tar")

l_obs        = request.Form("obs")

'Response.write "<script>alert('Estoy aca')</script>"

'Fec. Nac.
if len(l_fecnac) = 0 then
	l_fecnac = "null"
else 
	l_fecnac = cambiafecha(l_fecnac,"YMD",true)	
end if
'Est. Civ.
Select case l_estciv
	case "1"
		l_estcivnom = "SOLTERO/A"
	case "2"
		l_estcivnom = "CASADO/A"
	case "3"
		l_estcivnom = "DIVORCIADO/A"
	case "4"
		l_estcivnom = "SEPARADO/A"
	case "5"
		l_estcivnom = "VIUDO/A"
end select

'Localidad
'Select case l_loc
'	case "1"
'		l_locdes = "BAHIA BLANCA"
'	case "2"
'		l_locdes = "QUEQUEN"
'end select

'Provincia

'Tipo Doc
Select case l_tipdoc
	case "1"
		l_tipdocdes = "DNI"
	case "2"
		l_tipdocdes = "LC"
	case "3"
		l_tipdocdes = "LE"
end select

'Aporte Jubilatorio
Select case l_apojub
	case "1"
		l_apojubdes = "Reparto"
	case "2"
		l_apojubdes = "Capitalización"
end select

'Pad Fec Nac
if len(l_padfecnac) = 0 then
	l_padfecnac = "null"
else 
	l_padfecnac = cambiafecha(l_padfecnac,"YMD",true)	
end if

'Pad Vive
Select case l_padviv
	case "1"
		l_padviv = "Si"
	case "2"
		l_padviv = "No"
end select

'Pad Cargo
Select case l_padcar
	case "1"
		l_padcar = "Si"
	case "2"
		l_padcar = "No"
end select

'Pad Tipo Doc
Select case l_padtipdoc
	case "1"
		l_padtipdocdes = "DNI"
	case "2"
		l_padtipdocdes = "LC"
	case "3"
		l_padtipdocdes = "LE"
end select

'Mad Fec Nac
if len(l_madfecnac) = 0 then
	l_madfecnac = "null"
else 
	l_madfecnac = cambiafecha(l_madfecnac,"YMD",true)	
end if

'Mad Vive
Select case l_madviv
	case "1"
		l_madviv = "Si"
	case "2"
		l_madviv = "No"
end select

'Mad Cargo
Select case l_madcar
	case "1"
		l_madcar = "Si"
	case "2"
		l_madcar = "No"
end select

'Mad Tipo Doc
Select case l_madtipdoc
	case "1"
		l_madtipdocdes = "DNI"
	case "2"
		l_madtipdocdes = "LC"
	case "3"
		l_madtipdocdes = "LE"
end select

'Con Fec Nac
if len(l_confecnac) = 0 then
	l_confecnac = "null"
else 
	l_confecnac = cambiafecha(l_confecnac,"YMD",true)	
end if

'Con Vive
Select case l_conviv
	case "1"
		l_conviv = "Si"
	case "2"
		l_conviv = "No"
end select

'Con Cargo
Select case l_concar
	case "1"
		l_concar = "Si"
	case "2"
		l_concar = "No"
end select

'Con Tipo Doc
Select case l_contipdoc
	case "1"
		l_contipdocdes = "DNI"
	case "2"
		l_contipdocdes = "LC"
	case "3"
		l_contipdocdes = "LE"
end select

'Hi1 Fec Nac
if len(l_hi1fecnac) = 0 then
	l_hi1fecnac = "null"
else 
	l_hi1fecnac = cambiafecha(l_hi1fecnac,"YMD",true)	
end if

'Hi1 Vive
Select case l_hi1viv
	case "1"
		l_hi1viv = "Si"
	case "2"
		l_hi1viv = "No"
end select

'Hi1 Cargo
Select case l_hi1car
	case "1"
		l_hi1car = "Si"
	case "2"
		l_hi1car = "No"
end select

'Hi1 Tipo Doc
Select case l_hi1tipdoc
	case "1"
		l_hi1tipdocdes = "DNI"
	case "2"
		l_hi1tipdocdes = "LC"
	case "3"
		l_hi1tipdocdes = "LE"
end select

'Hi2 Fec Nac
if len(l_hi2fecnac) = 0 then
	l_hi2fecnac = "null"
else 
	l_hi2fecnac = cambiafecha(l_hi2fecnac,"YMD",true)	
end if

'Hi2 Vive
Select case l_hi2viv
	case "1"
		l_hi2viv = "Si"
	case "2"
		l_hi2viv = "No"
end select

'Hi2 Cargo
Select case l_hi2car
	case "1"
		l_hi2car = "Si"
	case "2"
		l_hi2car = "No"
end select

'Hi2 Tipo Doc
Select case l_hi2tipdoc
	case "1"
		l_hi2tipdocdes = "DNI"
	case "2"
		l_hi2tipdocdes = "LC"
	case "3"
		l_hi2tipdocdes = "LE"
end select

'Hi3 Fec Nac
if len(l_hi3fecnac) = 0 then
	l_hi3fecnac = "null"
else 
	l_hi3fecnac = cambiafecha(l_hi3fecnac,"YMD",true)
end if

'Hi3 Vive
Select case l_hi3viv
	case "1"
		l_hi3viv = "Si"
	case "2"
		l_hi3viv = "No"
end select

'Hi3 Cargo
Select case l_hi3car
	case "1"
		l_hi3car = "Si"
	case "2"
		l_hi3car = "No"
end select

'Hi3 Tipo Doc
Select case l_hi3tipdoc
	case "1"
		l_hi3tipdocdes = "DNI"
	case "2"
		l_hi3tipdocdes = "LC"
	case "3"
		l_hi3tipdocdes = "LE"
end select

'Hi4 Fec Nac
if len(l_hi4fecnac) = 0 then
	l_hi4fecnac = "null"
else 
	l_hi4fecnac = cambiafecha(l_hi4fecnac,"YMD",true)	
end if

'Hi4 Vive
Select case l_hi4viv
	case "1"
		l_hi4viv = "Si"
	case "2"
		l_hi4viv = "No"
end select

'Hi4 Cargo
Select case l_hi4car
	case "1"
		l_hi4car = "Si"
	case "2"
		l_hi4car = "No"
end select

'Hi4 Tipo Doc
Select case l_hi4tipdoc
	case "1"
		l_hi4tipdocdes = "DNI"
	case "2"
		l_hi4tipdocdes = "LC"
	case "3"
		l_hi4tipdocdes = "LE"
end select

'Hi5 Fec Nac
if len(l_hi5fecnac) = 0 then
	l_hi5fecnac = "null"
else 
	l_hi5fecnac = cambiafecha(l_hi5fecnac,"YMD",true)	
end if

'Hi5 Vive
Select case l_hi5viv
	case "1"
		l_hi5viv = "Si"
	case "2"
		l_hi5viv = "No"
end select

'Hi5 Cargo
Select case l_hi5car
	case "1"
		l_hi5car = "Si"
	case "2"
		l_hi5car = "No"
end select

'Hi5 Tipo Doc
Select case l_hi5tipdoc
	case "1"
		l_hi5tipdocdes = "DNI"
	case "2"
		l_hi5tipdocdes = "LC"
	case "3"
		l_hi5tipdocdes = "LE"
end select

'Hi6 Fec Nac
if len(l_hi6fecnac) = 0 then
	l_hi6fecnac = "null"
else 
	l_hi6fecnac = cambiafecha(l_hi6fecnac,"YMD",true)	
end if

'Hi6 Vive
Select case l_hi6viv
	case "1"
		l_hi6viv = "Si"
	case "2"
		l_hi6viv = "No"
end select

'Hi6 Cargo
Select case l_hi6car
	case "1"
		l_hi6car = "Si"
	case "2"
		l_hi6car = "No"
end select

'Hi6 Tipo Doc
Select case l_hi6tipdoc
	case "1"
		l_hi6tipdocdes = "DNI"
	case "2"
		l_hi6tipdocdes = "LC"
	case "3"
		l_hi6tipdocdes = "LE"
end select

'Pri desde
if len(l_prides) = 0 then
	l_prides = "null"
else 
	l_prides = cambiafecha(l_prides,"YMD",true)	
end if

'Pri hasta
if len(l_prihas) = 0 then
	l_prihas = "null"
else 
	l_prihas = cambiafecha(l_prihas,"YMD",true)	
end if

'Sec desde
if len(l_secdes) = 0 then
	l_secdes = "null"
else 
	l_secdes = cambiafecha(l_secdes,"YMD",true)	
end if

'Sec hasta
if len(l_sechas) = 0 then
	l_sechas = "null"
else 
	l_sechas = cambiafecha(l_sechas,"YMD",true)	
end if

'Ter desde
if len(l_terdes) = 0 then
	l_terdes = "null"
else 
	l_terdes = cambiafecha(l_terdes,"YMD",true)	
end if

'Ter hasta
if len(l_terhas) = 0 then
	l_terhas = "null"
else 
	l_terhas = cambiafecha(l_terhas,"YMD",true)	
end if

'Uni desde
if len(l_unides) = 0 then
	l_unides = "null"
else 
	l_unides = cambiafecha(l_unides,"YMD",true)	
end if

'Uni hasta
if len(l_unihas) = 0 then
	l_unihas = "null"
else 
	l_unihas = cambiafecha(l_unihas,"YMD",true)	
end if

'Pos desde
if len(l_posdes) = 0 then
	l_posdes = "null"
else 
	l_posdes = cambiafecha(l_posdes,"YMD",true)	
end if

'Pos hasta
if len(l_poshas) = 0 then
	l_poshas = "null"
else 
	l_poshas = cambiafecha(l_poshas,"YMD",true)	
end if

'Con desde
if len(l_condes) = 0 then
	l_condes = "null"
else 
	l_condes = cambiafecha(l_condes,"YMD",true)	
end if

'Con hasta
if len(l_conhas) = 0 then
	l_conhas = "null"
else 
	l_conhas = cambiafecha(l_conhas,"YMD",true)	
end if

'Idinom
Select case l_idinom
	case "0"
		l_idinom = ""
	case "1"
		l_idinom = "Inglés"
	case "2"
		l_idinom = "Frances"
	case "3"
		l_idinom = "Italiano"
	case "4"
		l_idinom = "Alemán"
	case "5"
		l_idinom = "Portugués"
end select

'Idilee
Select case l_idilee
	case "0"
		l_idilee = ""
	case "1"
		l_idilee = "Básico"
	case "2"
		l_idilee = "Intermedio"
	case "3"
		l_idilee = "Intermedio Avanzado"
	case "4"
		l_idilee = "Avanzado"
	case "5"
		l_idilee = "Bilingue"
end select

'Idihab
Select case l_idihab
	case "0"
		l_idihab = ""
	case "1"
		l_idihab = "Básico"
	case "2"
		l_idihab = "Intermedio"
	case "3"
		l_idihab = "Intermedio Avanzado"
	case "4"
		l_idihab = "Avanzado"
	case "5"
		l_idihab = "Bilingue"
end select

'Idiesc
Select case l_idiesc
	case "0"
		l_idiesc = ""
	case "1"
		l_idiesc = "Básico"
	case "2"
		l_idiesc = "Intermedio"
	case "3"
		l_idiesc = "Intermedio Avanzado"
	case "4"
		l_idiesc = "Avanzado"
	case "5"
		l_idiesc = "Bilingue"
end select

'Idi2nom
Select case l_idi2nom
	case "0"
		l_idi2nom = ""
	case "1"
		l_idi2nom = "Inglés"
	case "2"
		l_idi2nom = "Frances"
	case "3"
		l_idi2nom = "Italiano"
	case "4"
		l_idi2nom = "Alemán"
	case "5"
		l_idi2nom = "Portugués"
end select

'Idi2lee
Select case l_idi2lee
	case "0"
		l_idi2lee = ""
	case "1"
		l_idi2lee = "Básico"
	case "2"
		l_idi2lee = "Intermedio"
	case "3"
		l_idi2lee = "Intermedio Avanzado"
	case "4"
		l_idi2lee = "Avanzado"
	case "5"
		l_idi2lee = "Bilingue"
end select

'Idi2hab
Select case l_idi2hab
	case "0"
		l_idi2hab = ""
	case "1"
		l_idi2hab = "Básico"
	case "2"
		l_idi2hab = "Intermedio"
	case "3"
		l_idi2hab = "Intermedio Avanzado"
	case "4"
		l_idi2hab = "Avanzado"
	case "5"
		l_idi2hab = "Bilingue"
end select

'Idi2esc
Select case l_idi2esc
	case "0"
		l_idi2esc = ""
	case "1"
		l_idi2esc = "Básico"
	case "2"
		l_idi2esc = "Intermedio"
	case "3"
		l_idi2esc = "Intermedio Avanzado"
	case "4"
		l_idi2esc = "Avanzado"
	case "5"
		l_idi2esc = "Bilingue"
end select


'EmpAnt1 desde
if len(l_empant1des) = 0 then
	l_empant1des = "null"
else 
	l_empant1des = cambiafecha(l_empant1des,"YMD",true)	
end if

'Emp Ant1 hasta
if len(l_empant1has) = 0 then
	l_empant1has = "null"
else 
	l_empant1has = cambiafecha(l_empant1has,"YMD",true)	
end if

'EmpAnt2 desde
if len(l_empant2des) = 0 then
	l_empant2des = "null"
else 
	l_empant2des = cambiafecha(l_empant2des,"YMD",true)	
end if

'Emp Ant2 hasta
if len(l_empant2has) = 0 then
	l_empant2has = "null"
else 
	l_empant2has = cambiafecha(l_empant2has,"YMD",true)	
end if

'EmpAnt3 desde
if len(l_empant3des) = 0 then
	l_empant3des = "null"
else 
	l_empant3des = cambiafecha(l_empant3des,"YMD",true)	
end if

'Emp Ant3 hasta
if len(l_empant3has) = 0 then
	l_empant3has = "null"
else 
	l_empant3has = cambiafecha(l_empant3has,"YMD",true)	
end if

'Response.write "<script>alert(' y ahora aca Estoy aca')</script>"


	set l_cm = Server.CreateObject("ADODB.Command")
	
	if l_tipo = "M" then
	
		'Response.write "<script>alert(' antes del delete')</script>"
		
		l_sql = " DELETE FROM int_emp WHERE nrodoc = '" & l_nrodoc & "'"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	
	end if 
	
'	Response.write "<script>alert(' termine el delete')</script>"

	l_sql = " INSERT INTO int_emp (ape, nom , fecnac, nac, estciv, pue, cal, num, pis, dep, loc, pro, codpos, tel, cel, tipdoc, nrodoc, cuil, apojub, afjp, padape, padnom, padfecnac, padviv, padcar, padtipdoc, padnrodoc, madape, madnom, madfecnac, madviv, madcar, madtipdoc, madnrodoc , conape, connom, confecnac, conviv, concar, contipdoc, connrodoc, hi1ape, hi1nom, hi1fecnac, hi1viv, hi1car, hi1tipdoc, hi1nrodoc ,  hi2ape, hi2nom, hi2fecnac, hi2viv, hi2car, hi2tipdoc, hi2nrodoc , hi3ape, hi3nom , hi3fecnac, hi3viv  , hi3car, hi3tipdoc , hi3nrodoc , hi4ape, hi4nom, hi4fecnac, hi4viv, hi4car, hi4tipdoc, hi4nrodoc, hi5ape, hi5nom, hi5fecnac, hi5viv, hi5car, hi5tipdoc, hi5nrodoc ,  hi6ape, hi6nom , hi6fecnac, hi6viv , hi6car, hi6tipdoc, hi6nrodoc , priins, prides, prihas, pritit , secins, secdes, sechas, sectit , terins, terdes, terhas, tertit , uniins, unides, unihas, unitit , posins, posdes, poshas, postit , conins, condes, conhas, contit, idinom, idilee, idihab, idiesc, idi2nom, idi2lee, idi2hab, idi2esc,  empant1emp , empant1pue, empant1des, empant1has, empant1tar  , empant2emp , empant2pue, empant2des, empant2has, empant2tar , empant3emp , empant3pue, empant3des, empant3has, empant3tar, obs)"
	l_sql = l_sql & " VALUES ( '" & l_ape & "'"
	l_sql = l_sql & "        , '"  & l_nom & "'"
	l_sql = l_sql & "        ,  "  & l_fecnac
	l_sql = l_sql & "        , '"  & l_nac & "'"
	l_sql = l_sql & "        , '"  & l_estcivnom & "'"
	l_sql = l_sql & "        , '"  & l_pue & "'"
	l_sql = l_sql & "        , '"  & l_cal & "'"
	l_sql = l_sql & "        , '"  & l_num & "'"
	l_sql = l_sql & "        , '"  & l_pis & "'"
	l_sql = l_sql & "        , '"  & l_dep & "'"
	l_sql = l_sql & "        , '"  & l_loc & "'"
	l_sql = l_sql & "        , '"  & l_pro & "'"
	l_sql = l_sql & "        , '"  & l_codpos & "'"
	l_sql = l_sql & "        , '"  & l_tel & "'"
	l_sql = l_sql & "        , '"  & l_cel & "'"
	l_sql = l_sql & "        , '"  & l_tipdocdes & "'"
	l_sql = l_sql & "        , '"  & l_nrodoc & "'"
	l_sql = l_sql & "        , '"  & l_cuil & "'"
	l_sql = l_sql & "        , '"  & l_apojubdes & "'"
	l_sql = l_sql & "        , '"  & l_afjp & "'"
	
	l_sql = l_sql & "        , '"  & l_padape & "'"
	l_sql = l_sql & "        , '"  & l_padnom & "'"
	l_sql = l_sql & "        ,  "  & l_padfecnac
	l_sql = l_sql & "        , '"  & l_padviv & "'"	
	l_sql = l_sql & "        , '"  & l_padcar & "'"	
	l_sql = l_sql & "        , '"  & l_padtipdocdes & "'"
	l_sql = l_sql & "        , '"  & l_padnrodoc & "'"	
	
	l_sql = l_sql & "        , '"  & l_madape & "'"
	l_sql = l_sql & "        , '"  & l_madnom & "'"
	l_sql = l_sql & "        ,  "  & l_madfecnac
	l_sql = l_sql & "        , '"  & l_madviv & "'"	
	l_sql = l_sql & "        , '"  & l_madcar & "'"	
	l_sql = l_sql & "        , '"  & l_madtipdocdes & "'"
	l_sql = l_sql & "        , '"  & l_madnrodoc & "'"
	
	l_sql = l_sql & "        , '"  & l_conape & "'"
	l_sql = l_sql & "        , '"  & l_connom & "'"
	l_sql = l_sql & "        ,  "  & l_confecnac
	l_sql = l_sql & "        , '"  & l_conviv & "'"	
	l_sql = l_sql & "        , '"  & l_concar & "'"	
	l_sql = l_sql & "        , '"  & l_contipdocdes & "'"
	l_sql = l_sql & "        , '"  & l_connrodoc & "'"
	
	l_sql = l_sql & "        , '"  & l_hi1ape & "'"
	l_sql = l_sql & "        , '"  & l_hi1nom & "'"
	l_sql = l_sql & "        ,  "  & l_hi1fecnac
	l_sql = l_sql & "        , '"  & l_hi1viv & "'"	
	l_sql = l_sql & "        , '"  & l_hi1car & "'"	
	l_sql = l_sql & "        , '"  & l_hi1tipdocdes & "'"
	l_sql = l_sql & "        , '"  & l_hi1nrodoc & "'"	
	
	l_sql = l_sql & "        , '"  & l_hi2ape & "'"
	l_sql = l_sql & "        , '"  & l_hi2nom & "'"
	l_sql = l_sql & "        ,  "  & l_hi2fecnac
	l_sql = l_sql & "        , '"  & l_hi2viv & "'"	
	l_sql = l_sql & "        , '"  & l_hi2car & "'"	
	l_sql = l_sql & "        , '"  & l_hi2tipdocdes & "'"
	l_sql = l_sql & "        , '"  & l_hi2nrodoc & "'"		
	
	l_sql = l_sql & "        , '"  & l_hi3ape & "'"
	l_sql = l_sql & "        , '"  & l_hi3nom & "'"
	l_sql = l_sql & "        ,  "  & l_hi3fecnac
	l_sql = l_sql & "        , '"  & l_hi3viv & "'"	
	l_sql = l_sql & "        , '"  & l_hi3car & "'"	
	l_sql = l_sql & "        , '"  & l_hi3tipdocdes & "'"
	l_sql = l_sql & "        , '"  & l_hi3nrodoc & "'"
	
	l_sql = l_sql & "        , '"  & l_hi4ape & "'"
	l_sql = l_sql & "        , '"  & l_hi4nom & "'"
	l_sql = l_sql & "        ,  "  & l_hi4fecnac
	l_sql = l_sql & "        , '"  & l_hi4viv & "'"	
	l_sql = l_sql & "        , '"  & l_hi4car & "'"	
	l_sql = l_sql & "        , '"  & l_hi4tipdocdes & "'"
	l_sql = l_sql & "        , '"  & l_hi4nrodoc & "'"	
	
	l_sql = l_sql & "        , '"  & l_hi5ape & "'"
	l_sql = l_sql & "        , '"  & l_hi5nom & "'"
	l_sql = l_sql & "        ,  "  & l_hi5fecnac
	l_sql = l_sql & "        , '"  & l_hi5viv & "'"	
	l_sql = l_sql & "        , '"  & l_hi5car & "'"	
	l_sql = l_sql & "        , '"  & l_hi5tipdocdes & "'"
	l_sql = l_sql & "        , '"  & l_hi5nrodoc & "'"	
	
	l_sql = l_sql & "        , '"  & l_hi6ape & "'"
	l_sql = l_sql & "        , '"  & l_hi6nom & "'"
	l_sql = l_sql & "        ,  "  & l_hi6fecnac
	l_sql = l_sql & "        , '"  & l_hi6viv & "'"	
	l_sql = l_sql & "        , '"  & l_hi6car & "'"	
	l_sql = l_sql & "        , '"  & l_hi6tipdocdes & "'"
	l_sql = l_sql & "        , '"  & l_hi6nrodoc & "'"	
	
	l_sql = l_sql & "        , '"  & l_priins & "'"
	l_sql = l_sql & "        ,  "  & l_prides
	l_sql = l_sql & "        ,  "  & l_prihas
	l_sql = l_sql & "        , '"  & l_pritit & "'"
	
	l_sql = l_sql & "        , '"  & l_secins & "'"
	l_sql = l_sql & "        ,  "  & l_secdes
	l_sql = l_sql & "        ,  "  & l_sechas
	l_sql = l_sql & "        , '"  & l_sectit & "'"	
	
	l_sql = l_sql & "        , '"  & l_terins & "'"
	l_sql = l_sql & "        ,  "  & l_terdes
	l_sql = l_sql & "        ,  "  & l_terhas
	l_sql = l_sql & "        , '"  & l_tertit & "'"
	
	l_sql = l_sql & "        , '"  & l_uniins & "'"
	l_sql = l_sql & "        ,  "  & l_unides
	l_sql = l_sql & "        ,  "  & l_unihas
	l_sql = l_sql & "        , '"  & l_unitit & "'"	
	
	l_sql = l_sql & "        , '"  & l_posins & "'"
	l_sql = l_sql & "        ,  "  & l_posdes
	l_sql = l_sql & "        ,  "  & l_poshas
	l_sql = l_sql & "        , '"  & l_postit & "'"
	
	l_sql = l_sql & "        , '"  & l_conins & "'"
	l_sql = l_sql & "        ,  "  & l_condes
	l_sql = l_sql & "        ,  "  & l_conhas
	l_sql = l_sql & "        , '"  & l_contit & "'"	
	
	l_sql = l_sql & "        , '"  & l_idinom & "'"
	l_sql = l_sql & "        , '"  & l_idilee & "'"
	l_sql = l_sql & "        , '"  & l_idihab & "'"
	l_sql = l_sql & "        , '"  & l_idiesc & "'"
	
	l_sql = l_sql & "        , '"  & l_idi2nom & "'"
	l_sql = l_sql & "        , '"  & l_idi2lee & "'"
	l_sql = l_sql & "        , '"  & l_idi2hab & "'"
	l_sql = l_sql & "        , '"  & l_idi2esc & "'"
	
	l_sql = l_sql & "        , '"  & l_empant1emp & "'"
	l_sql = l_sql & "        , '"  & l_empant1pue & "'"	
	l_sql = l_sql & "        ,  "  & l_empant1des
	l_sql = l_sql & "        ,  "  & l_empant1has
	l_sql = l_sql & "        , '"  & l_empant1tar & "'"

	l_sql = l_sql & "        , '"  & l_empant2emp & "'"
	l_sql = l_sql & "        , '"  & l_empant2pue & "'"	
	l_sql = l_sql & "        ,  "  & l_empant2des
	l_sql = l_sql & "        ,  "  & l_empant2has
	l_sql = l_sql & "        , '"  & l_empant2tar & "'"	
	
	l_sql = l_sql & "        , '"  & l_empant3emp & "'"
	l_sql = l_sql & "        , '"  & l_empant3pue & "'"	
	l_sql = l_sql & "        ,  "  & l_empant3des
	l_sql = l_sql & "        ,  "  & l_empant3has
	l_sql = l_sql & "        , '"  & l_empant3tar & "'"	

	l_sql = l_sql & "        , '"  & l_obs & "'"
	
	l_sql = l_sql & " ) "

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "<script>alert('Operación realizada. Los datos han sido almacenados. \n                 Gracias por su colaboración.');window.parent.location='datos_personal_con_02.asp?id=0'</script>"
%>

