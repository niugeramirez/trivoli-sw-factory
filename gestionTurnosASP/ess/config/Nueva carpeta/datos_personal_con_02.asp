<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: datos_personal_con_02.asp
'Descripción: Ingreso de datos de los Empleados
'Autor : Raul Chinestra
'Fecha: 16/07/2008

'ADO
Dim l_sql
Dim l_rs
Dim l_tipo

Dim l_empnro
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

on error goto 0

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ingreso de datos del Empleado</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>

function Validar(){

if (stringValido(document.datos.id.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.id.focus();
	return;
}
if (document.datos.id.value == "") {
	alert("Debe ingresar un Nro. de Documento.");
	document.datos.id.focus();
	return;
}

document.location = "datos_personal_con_02.asp?id="+document.datos.id.value;

//valido();

}

function Asignar_Provincia(){

	//alert(document.datos.loc.value);
	
	document.ifrmpro.location = "provincias_con_00.asp?locdes="+document.datos.loc.value;

}


function Validar_Formulario(){


// ********************************************************************************************
// Verificar que se ingresen caracteres Válidos

if (stringValido(document.datos.ape.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.ape.focus();
	return;
}
if (stringValido(document.datos.nom.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.nom.focus();
	return;
}
if (stringValido(document.datos.nac.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.nac.focus();
	return;
}
if (stringValido(document.datos.pue.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.pue.focus();
	return;
}
if (stringValido(document.datos.cal.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.cal.focus();
	return;
}
if (stringValido(document.datos.num.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.num.focus();
	return;
}
if (stringValido(document.datos.pis.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.pis.focus();
	return;
}
if (stringValido(document.datos.dep.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.dep.focus();
	return;
}
if (stringValido(document.datos.codpos.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.codpos.focus();
	return;
}
if (stringValido(document.datos.tel.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.tel.focus();
	return;
}
if (stringValido(document.datos.cel.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.cel.focus();
	return;
}
if (stringValido(document.datos.nrodoc.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.nrodoc.focus();
	return;
}
if (stringValido(document.datos.cuil.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.cuil.focus();
	return;
}
if (stringValido(document.datos.afjp.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.afjp.focus();
	return;
}
if (stringValido(document.datos.padape.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.padape.focus();
	return;
}
if (stringValido(document.datos.padnom.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.padnom.focus();
	return;
}
if (stringValido(document.datos.padnrodoc.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.padnrodoc.focus();
	return;
}
if (stringValido(document.datos.madape.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.madape.focus();
	return;
}
if (stringValido(document.datos.madnom.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.madnom.focus();
	return;
}
if (stringValido(document.datos.madnrodoc.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.madnrodoc.focus();
	return;
}
if (stringValido(document.datos.conape.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.conape.focus();
	return;
}
if (stringValido(document.datos.connom.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.connom.focus();
	return;
}
if (stringValido(document.datos.connrodoc.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.connrodoc.focus();
	return;
}
if (stringValido(document.datos.hi1ape.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi1ape.focus();
	return;
}
if (stringValido(document.datos.hi1nom.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi1nom.focus();
	return;
}
if (stringValido(document.datos.hi1nrodoc.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi1nrodoc.focus();
	return;
}
if (stringValido(document.datos.hi2ape.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi2ape.focus();
	return;
}
if (stringValido(document.datos.hi2nom.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi2nom.focus();
	return;
}
if (stringValido(document.datos.hi2nrodoc.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi2nrodoc.focus();
	return;
}
if (stringValido(document.datos.hi3ape.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi3ape.focus();
	return;
}
if (stringValido(document.datos.hi3nom.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi3nom.focus();
	return;
}
if (stringValido(document.datos.hi3nrodoc.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi3nrodoc.focus();
	return;
}
if (stringValido(document.datos.hi4ape.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi4ape.focus();
	return;
}
if (stringValido(document.datos.hi4nom.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi4nom.focus();
	return;
}
if (stringValido(document.datos.hi4nrodoc.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi4nrodoc.focus();
	return;
}
if (stringValido(document.datos.hi5ape.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi5ape.focus();
	return;
}
if (stringValido(document.datos.hi5nom.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi5nom.focus();
	return;
}
if (stringValido(document.datos.hi5nrodoc.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi5nrodoc.focus();
	return;
}
if (stringValido(document.datos.hi6ape.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi6ape.focus();
	return;
}
if (stringValido(document.datos.hi6nom.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi6nom.focus();
	return;
}
if (stringValido(document.datos.hi6nrodoc.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.hi6nrodoc.focus();
	return;
}
if (stringValido(document.datos.priins.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.priins.focus();
	return;
}
if (stringValido(document.datos.pritit.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.pritit.focus();
	return;
}
if (stringValido(document.datos.secins.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.secins.focus();
	return;
}
if (stringValido(document.datos.sectit.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.sectit.focus();
	return;
}
if (stringValido(document.datos.terins.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.terins.focus();
	return;
}
if (stringValido(document.datos.tertit.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.tertit.focus();
	return;
}
if (stringValido(document.datos.uniins.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.uniins.focus();
	return;
}
if (stringValido(document.datos.unitit.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.unitit.focus();
	return;
}
if (stringValido(document.datos.posins.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.posins.focus();
	return;
}
if (stringValido(document.datos.postit.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.postit.focus();
	return;
}
if (stringValido(document.datos.conins.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.conins.focus();
	return;
}
if (stringValido(document.datos.contit.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.contit.focus();
	return;
}
if (stringValido(document.datos.empant1emp.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.empant1emp.focus();
	return;
}
if (stringValido(document.datos.empant1pue.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.empant1pue.focus();
	return;
}
if (stringValido(document.datos.empant1tar.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.empant1tar.focus();
	return;
}
if (stringValido(document.datos.empant2emp.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.empant2emp.focus();
	return;
}
if (stringValido(document.datos.empant2pue.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.empant2pue.focus();
	return;
}
if (stringValido(document.datos.empant2tar.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.empant2tar.focus();
	return;
}
if (stringValido(document.datos.empant3emp.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.empant3emp.focus();
	return;
}
if (stringValido(document.datos.empant3pue.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.empant3pue.focus();
	return;
}
if (stringValido(document.datos.empant3tar.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.empant3tar.focus();
	return;
}
if (stringValido(document.datos.obs.value) == false) {
	alert("Se ingresaron Caracteres no Válidos, verificar los mismos.");
	document.datos.obs.focus();
	return;
}

// ********************************************************************************************

if (document.datos.ape.value == ""){
	alert("Debe ingresar el Apellido.");
	document.datos.ape.focus();
	return;
}
if (document.datos.nom.value == ""){
	alert("Debe ingresar el Nombre.");
	document.datos.nom.focus();
	return;
}
if (document.datos.fecnac.value == 0){
	alert("Debe ingresar la Fecha de Nacimiento.");
	document.datos.fecnac.focus();
	return;
}
if ((document.datos.fecnac.value != "")&&(!validarfecha(document.datos.fecnac))){
	 document.datos.fecnac.focus();
	 return;
}
if (document.datos.nac.value == ""){
	alert("Debe ingresar la Nacionalidad.");
	document.datos.nac.focus();
	return;
}
if (document.datos.estciv.value == 0){
	alert("Debe ingresar el Estado Civil.");
	document.datos.estciv.focus();
	return;
}
if (document.datos.pue.value == ""){
	alert("Debe ingresar el Puesto.");
	document.datos.pue.focus();
	return;
}
if (document.datos.cal.value == ""){
	alert("Debe ingresar la Calle.");
	document.datos.cal.focus();
	return;
}
if (document.datos.num.value == ""){
	alert("Debe ingresar el Número.");
	document.datos.num.focus();
	return;
}
if (document.datos.loc.value == 0){
	alert("Debe ingresar la Localidad.");
	document.datos.loc.focus();
	return;
}
if (document.ifrmpro.datos.pronro.value == 0){
	alert("Debe ingresar la Provincia.");
	document.ifrmpro.datos.pronro.focus();
	return;
}
if (document.datos.codpos.value == ""){
	alert("Debe ingresar el Código Postal.");
	document.datos.codpos.focus();
	return;
}
if (document.datos.tipdoc.value == 0){
	alert("Debe ingresar el Tipo de Documento.");
	document.datos.tipdoc.focus();
	return;
}
if (document.datos.nrodoc.value == ""){
	alert("Debe ingresar el Nro. de Documento.");
	document.datos.nrodoc.focus();
	return;
}
if (document.datos.apojub.value == 0){
	alert("Debe ingresar el Aporte Jubilatorio.");
	document.datos.apojub.focus();
	return;
}


if ((document.datos.padfecnac.value != "")&&(!validarfecha(document.datos.padfecnac))){
	 document.datos.padfecnac.focus();
	 return;
}
if ((document.datos.madfecnac.value != "")&&(!validarfecha(document.datos.madfecnac))){
	 document.datos.madfecnac.focus();
	 return;
}
if ((document.datos.confecnac.value != "")&&(!validarfecha(document.datos.confecnac))){
	 document.datos.confecnac.focus();
	 return;
}
if ((document.datos.hi1fecnac.value != "")&&(!validarfecha(document.datos.hi1fecnac))){
	 document.datos.hi1fecnac.focus();
	 return;
}
if ((document.datos.hi2fecnac.value != "")&&(!validarfecha(document.datos.hi2fecnac))){
	 document.datos.hi2fecnac.focus();
	 return;
}
if ((document.datos.hi3fecnac.value != "")&&(!validarfecha(document.datos.hi3fecnac))){
	 document.datos.hi3fecnac.focus();
	 return;
}
if ((document.datos.hi4fecnac.value != "")&&(!validarfecha(document.datos.hi4fecnac))){
	 document.datos.hi4fecnac.focus();
	 return;
}
if ((document.datos.hi5fecnac.value != "")&&(!validarfecha(document.datos.hi5fecnac))){
	 document.datos.hi5fecnac.focus();
	 return;
}
if ((document.datos.hi6fecnac.value != "")&&(!validarfecha(document.datos.hi6fecnac))){
	 document.datos.hi6fecnac.focus();
	 return;
}
if ((document.datos.prides.value != "")&&(!validarfecha(document.datos.prides))){
	 document.datos.prides.focus();
	 return;
}
if ((document.datos.prihas.value != "")&&(!validarfecha(document.datos.prihas))){
	 document.datos.prihas.focus();
	 return;
}
if ((document.datos.prides.value != "")&&(document.datos.prihas.value != "")) {
	if (!(menorque(document.datos.prides.value,document.datos.prihas.value))) {
		alert("La Fecha Desde debe ser menor que la Fecha Hasta.");	
		document.datos.prides.focus();
		return;	
	}	
}		
if ((document.datos.secdes.value != "")&&(!validarfecha(document.datos.secdes))){
	 document.datos.secdes.focus();
	 return;
}
if ((document.datos.sechas.value != "")&&(!validarfecha(document.datos.sechas))){
	 document.datos.sechas.focus();
	 return;
}
if ((document.datos.secdes.value != "")&&(document.datos.sechas.value != "")) {
	if (!(menorque(document.datos.secdes.value,document.datos.sechas.value))) {
		alert("La Fecha Desde debe ser menor que la Fecha Hasta.");	
		document.datos.secdes.focus();
		return;	
	}	
}		
if ((document.datos.terdes.value != "")&&(!validarfecha(document.datos.terdes))){
	 document.datos.terdes.focus();
	 return;
}
if ((document.datos.terhas.value != "")&&(!validarfecha(document.datos.terhas))){
	 document.datos.terhas.focus();
	 return;
}
if ((document.datos.terdes.value != "")&&(document.datos.terhas.value != "")) {
	if (!(menorque(document.datos.terdes.value,document.datos.terhas.value))) {
		alert("La Fecha Desde debe ser menor que la Fecha Hasta.");	
		document.datos.terdes.focus();
		return;	
	}	
}		
if ((document.datos.unides.value != "")&&(!validarfecha(document.datos.unides))){
	 document.datos.unides.focus();
	 return;
}
if ((document.datos.unihas.value != "")&&(!validarfecha(document.datos.unihas))){
	 document.datos.unihas.focus();
	 return;
}
if ((document.datos.unides.value != "")&&(document.datos.unihas.value != "")) {
	if (!(menorque(document.datos.unides.value,document.datos.unihas.value))) {
		alert("La Fecha Desde debe ser menor que la Fecha Hasta.");	
		document.datos.unides.focus();
		return;	
	}	
}		
if ((document.datos.posdes.value != "")&&(!validarfecha(document.datos.posdes))){
	 document.datos.posdes.focus();
	 return;
}
if ((document.datos.poshas.value != "")&&(!validarfecha(document.datos.poshas))){
	 document.datos.poshas.focus();
	 return;
}
if ((document.datos.posdes.value != "")&&(document.datos.poshas.value != "")) {
	if (!(menorque(document.datos.posdes.value,document.datos.poshas.value))) {
		alert("La Fecha Desde debe ser menor que la Fecha Hasta.");	
		document.datos.posdes.focus();
		return;	
	}	
}		
if ((document.datos.condes.value != "")&&(!validarfecha(document.datos.condes))){
	 document.datos.condes.focus();
	 return;
}
if ((document.datos.conhas.value != "")&&(!validarfecha(document.datos.conhas))){
	 document.datos.conhas.focus();
	 return;
}
if ((document.datos.condes.value != "")&&(document.datos.conhas.value != "")) {
	if (!(menorque(document.datos.condes.value,document.datos.conhas.value))) {
		alert("La Fecha Desde debe ser menor que la Fecha Hasta.");	
		document.datos.condes.focus();
		return;	
	}	
}		
if ((document.datos.empant1des.value != "")&&(!validarfecha(document.datos.empant1des))){
	 document.datos.empant1des.focus();
	 return;
}
if ((document.datos.empant1has.value != "")&&(!validarfecha(document.datos.empant1has))){
	 document.datos.empant1has.focus();
	 return;
}
if ((document.datos.empant1des.value != "")&&(document.datos.empant1has.value != "")) {
	if (!(menorque(document.datos.empant1des.value,document.datos.empant1has.value))) {
		alert("La Fecha Desde debe ser menor que la Fecha Hasta.");	
		document.datos.empant1des.focus();
		return;	
	}	
}
if ((document.datos.empant2des.value != "")&&(!validarfecha(document.datos.empant2des))){
	 document.datos.empant2des.focus();
	 return;
}
if ((document.datos.empant2has.value != "")&&(!validarfecha(document.datos.empant2has))){
	 document.datos.empant2has.focus();
	 return;
}
if ((document.datos.empant2des.value != "")&&(document.datos.empant2has.value != "")) {
	if (!(menorque(document.datos.empant2des.value,document.datos.empant2has.value))) {
		alert("La Fecha Desde debe ser menor que la Fecha Hasta.");	
		document.datos.empant2des.focus();
		return;	
	}	
}
if ((document.datos.empant3des.value != "")&&(!validarfecha(document.datos.empant3des))){
	 document.datos.empant3des.focus();
	 return;
}
if ((document.datos.empant3has.value != "")&&(!validarfecha(document.datos.empant3has))){
	 document.datos.empant3has.focus();
	 return;
}
if ((document.datos.empant3des.value != "")&&(document.datos.empant3has.value != "")) {
	if (!(menorque(document.datos.empant3des.value,document.datos.empant3has.value))) {
		alert("La Fecha Desde debe ser menor que la Fecha Hasta.");	
		document.datos.empant3des.focus();
		return;	
	}	
}

if (document.datos.obs.value.length > 500) 
{
  	alert("La Observaciones no pueden superar los 500 Caracteres.");
	document.datos.obs.focus();
	return;
}

valido();


/*

// Purchases - Sales
if (document.datos.conpursal.value == "0"){
	alert("Debe ingresar Purchase (P) - Sale (S).");
	document.datos.conpursal.focus();
	return;
}
// Date
if (document.datos.confec.value == 0){
	alert("Debe ingresar el Date.");
	document.datos.confec.focus();
	return;
}
if ((document.datos.confec.value != "")&&(!validarfecha(document.datos.confec))){
	 document.datos.confec.focus();
	 return;
}
// Office
if (document.datos.offnro.value == 0){
	alert("Debe ingresar la Office asociada al Contrato.");
	document.datos.offnro.focus();
	return;
} 
// Ctr Number
if (document.datos.ctrnum.value == 0){
	alert("Debe ingresar el Ctr Number asociado al Contrato.");
	document.datos.ctrnum.focus();
	return;
}
// Company
if (document.datos.comnro.value == 0){
	alert("Debe ingresar la Company asociada al Contrato.");
	document.datos.comnro.focus();
	return;
}
	}
	
	
	
	if (!(menorque(document.datos.conarrini.value,document.datos.conarrfin.value))) {
		alert("La Fecha de Arrival Inicio debe ser menor o igual que la Fecha Arrival Final.");	
		document.datos.conarrini.focus();
		return;	
	}	
	
}
// Extension
if (document.datos.conext.value == ""){
	alert("Debe Ingresar un Valor en Extension");
	document.datos.conext.focus();
	return;
}		
if (isNaN(document.datos.conext.value)){
	alert("Debe Ingresar un Valor Numérico en Extension");
	document.datos.conext.focus();
	return;
}		
// Ports
if (document.datos.pornro.value == 0){
	alert("Debe ingresar el Port asociado al Contrato.");
	document.datos.pornro.focus();
	return;
}
// Berth
if (document.datos.bernro.value == 0){
	alert("Debe ingresar el Berth asociado al Contrato.");
	document.ifrmberth.datos.bernro.focus();
	return;
}
// Terms
if (document.datos.ternro.value == 0){
	alert("Debe ingresar el Term asociado al Contrato.");
	document.datos.ternro.focus();
	return;
}
// Crop Year
if (Trim(document.datos.croyea.value) == ""){
	alert("Debe ingresar el Crop Year.");
	document.datos.croyea.focus();
	return;
}
if (isNaN(document.datos.croyea.value)){
	alert("Debe Ingresar un Valor Numérico en el Crop Year");
	document.datos.croyea.focus();
	return;
}
// Preadvice
if (isNaN(document.datos.conpreadv.value)){
	alert("Debe Ingresar un Valor Numérico en Preadvice");
	document.datos.conpreadv.focus();
	return;
}
// Weight
if (document.datos.conwei.value == 0){
	alert("Debe ingresar El Weigth asociado al Contrato.");
	document.datos.conwei.focus();
	return;
}
// Quality
if (document.datos.conqua.value == 0){
	alert("Debe ingresar la Quality asociada al Contrato.");
	document.datos.conqua.focus();
	return;
}
// Commision
if (isNaN(document.datos.concom.value)){
	alert("Debe Ingresar un Valor Numérico en Commision");
	document.datos.concom.focus();
	return;
}
if (document.datos.concom.value == ""){
	document.datos.concom.value = 0;
}
// Qual Allowance
if (isNaN(document.datos.conquaall.value)){
	alert("Debe Ingresar un Valor Numérico en Qual Allowance");
	document.datos.conquaall.focus();
	return;
}
if (document.datos.conquaall.value == ""){
	document.datos.conquaall.value = 0;
}
// Other Cost
if (isNaN(document.datos.conothcos.value)){
	alert("Debe Ingresar un Valor Numérico en Other Cost");
	document.datos.conothcos.focus();
	return;
}
if (document.datos.conothcos.value == ""){
	document.datos.conothcos.value = 0;
}
// Fob Parity
if (isNaN(document.datos.confobpar.value)){
	alert("Debe Ingresar un Valor Numérico en Fob Parity");
	document.datos.confobpar.focus();
	return;
}
if (document.datos.confobpar.value == ""){
	document.datos.confobpar.value = 0;
}
// Freigh Rate
if (isNaN(document.datos.confrerat.value)){
	alert("Debe Ingresar un Valor Numérico en Freight Rate");
	document.datos.confrerat.focus();
	return;
}
if (document.datos.confrerat.value == ""){
	document.datos.confrerat.value = 0;
}
// Finance
if (isNaN(document.datos.confin.value)){
	alert("Debe Ingresar un Valor Numérico en Finance");
	document.datos.confin.focus();
	return;
}
if (document.datos.confin.value == ""){
	document.datos.confin.value = 0;
}
// Fog 
if (isNaN(document.datos.confog.value)){
	alert("Debe Ingresar un Valor Numérico en Fog");
	document.datos.confog.focus();
	return;
}
if (document.datos.confog.value == ""){
	document.datos.confog.value = 0;
}

Actualizar_Price();

/*
var d=document.datos;
document.valida.location = "countries_con_06.asp?tipo=<%'= l_tipo%>&counro="+document.datos.counro.value + "&coudes="+document.datos.coudes.value;
*/
//valido();



}

function valido(){
	document.datos.submit();
}


function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}


</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'response.write "raul" & request.querystring("id")

l_empnro = request.querystring("id")
'response.write "999999999999999"
l_sql = "SELECT  * "
l_sql = l_sql & " FROM int_emp "
l_sql  = l_sql  & " WHERE nrodoc = '" & l_empnro & "'"
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_tipo = "M"
    '************************************************
	l_ape = l_rs("ape")
	l_nom = l_rs("nom")
	l_fecnac = l_rs("fecnac")
	l_nac = l_rs("nac")
	l_estciv = l_rs("estciv")
	select case l_estciv 
		case "SOLTERO/A"
			 l_estciv = 1
		case "CASADO/A"
			 l_estciv = 2
		case "DIVORCIADO/A"
		 	 l_estciv = 3
		case "SEPARADO/A"
			 l_estciv = 4
		case "VIUDO/A"
	  		 l_estciv  = 5
	end select
	l_pue = l_rs("pue")
    '************************************************	
	l_cal = l_rs("cal")
	l_num = l_rs("num")
	l_pis = l_rs("pis")
	l_dep = l_rs("dep")
	l_loc = l_rs("loc")
'	Select case l_loc
'		case "BAHIA BLANCA"
'			l_loc = 1
'		case "QUEQUEN"
'			l_loc = 2
'	end select
	l_pro = l_rs("pro")
	l_codpos = l_rs("codpos")
	l_tel = l_rs("tel")
	l_cel = l_rs("cel")
    '************************************************
	l_tipdoc = l_rs("tipdoc")
	Select case l_tipdoc
		case "DNI"
			l_tipdoc = 1
		case "LC"
			l_tipdoc = 2
		case "LE"
			l_tipdoc = 3
	end select
	l_nrodoc = l_rs("nrodoc")
	l_cuil = l_rs("cuil")
	l_apojub = l_rs("apojub")
	Select case l_apojub
		case "Reparto"
			l_apojub = 1
		case "Capitalización"
			l_apojub = 2
	end select
	l_afjp = l_rs("afjp")
    '************************************************
	l_padape = l_rs("padape")
	l_padnom = l_rs("padnom")
	l_padfecnac = l_rs("padfecnac")
	l_padviv = l_rs("padviv")
	Select case l_padviv
		case "Si"
			l_padviv = 1
		case "No"
			l_padviv = 2
	end select	
	l_padcar = l_rs("padcar")
	Select case l_padcar
		case "Si"
			l_padcar = 1
		case "No"
			l_padcar = 2
	end select	
	l_padtipdoc = l_rs("padtipdoc")
	Select case l_padtipdoc
		case "DNI"
			l_padtipdoc = 1
		case "LC"
			l_padtipdoc = 2
		case "LE"
			l_padtipdoc = 3
	end select	
	l_padnrodoc = l_rs("padnrodoc")
    '************************************************
	l_madape = l_rs("madape")
	l_madnom = l_rs("madnom")
	l_madfecnac = l_rs("madfecnac")
	l_madviv = l_rs("madviv")
	Select case l_madviv
		case "Si"
			l_madviv = 1
		case "No"
			l_madviv = 2
	end select	
	l_madcar = l_rs("madcar")
	Select case l_madcar
		case "Si"
			l_madcar = 1
		case "No"
			l_madcar = 2
	end select	
	l_madtipdoc = l_rs("madtipdoc")
	Select case l_madtipdoc
		case "DNI"
			l_madtipdoc = 1
		case "LC"
			l_madtipdoc = 2
		case "LE"
			l_madtipdoc = 3
	end select
	l_madnrodoc = l_rs("madnrodoc")
   '************************************************
	l_conape = l_rs("conape")
	l_connom = l_rs("connom")
	l_confecnac = l_rs("confecnac")
	l_conviv = l_rs("conviv")
	Select case l_conviv
		case "Si"
			l_conviv = 1
		case "No"
			l_conviv = 2
	end select	
	l_concar = l_rs("concar")
	Select case l_concar
		case "Si"
			l_concar = 1
		case "No"
			l_concar = 2
	end select		
	l_contipdoc = l_rs("contipdoc")
	Select case l_contipdoc
		case "DNI"
			l_contipdoc = 1
		case "LC"
			l_contipdoc = 2
		case "LE"
			l_contipdoc = 3
	end select	
	l_connrodoc = l_rs("connrodoc")
   '************************************************
	l_hi1ape = l_rs("hi1ape")
	l_hi1nom = l_rs("hi1nom")
	l_hi1fecnac = l_rs("hi1fecnac")
	l_hi1viv = l_rs("hi1viv")
	Select case l_hi1viv
		case "Si"
			l_hi1viv = 1
		case "No"
			l_hi1viv = 2
	end select	
	l_hi1car = l_rs("hi1car")
	Select case l_hi1car
		case "Si"
			l_hi1car = 1
		case "No"
			l_hi1car = 2
	end select			
	l_hi1tipdoc = l_rs("hi1tipdoc")
	Select case l_hi1tipdoc
		case "DNI"
			l_hi1tipdoc = 1
		case "LC"
			l_hi1tipdoc = 2
		case "LE"
			l_hi1tipdoc = 3
	end select		
	l_hi1nrodoc = l_rs("hi1nrodoc")
   '************************************************
	l_hi2ape = l_rs("hi2ape")
	l_hi2nom = l_rs("hi2nom")
	l_hi2fecnac = l_rs("hi2fecnac") 
	l_hi2viv = l_rs("hi2viv")
	Select case l_hi2viv
		case "Si"
			l_hi2viv = 1
		case "No"
			l_hi2viv = 2
	end select	
	l_hi2car = l_rs("hi2car")
	Select case l_hi2car
		case "Si"
			l_hi2car = 1
		case "No"
			l_hi2car = 2
	end select	
	l_hi2tipdoc = l_rs("hi2tipdoc")
	Select case l_hi2tipdoc
		case "DNI"
			l_hi2tipdoc = 1
		case "LC"
			l_hi2tipdoc = 2
		case "LE"
			l_hi2tipdoc = 3
	end select	
	l_hi2nrodoc = l_rs("hi2nrodoc")
   '************************************************
	l_hi3ape = l_rs("hi3ape")
	l_hi3nom = l_rs("hi3nom")
	l_hi3fecnac = l_rs("hi3fecnac")
	l_hi3viv = l_rs("hi3viv")
	Select case l_hi3viv
		case "Si"
			l_hi3viv = 1
		case "No"
			l_hi3viv = 2
	end select	
	l_hi3car = l_rs("hi3car")
	Select case l_hi3car
		case "Si"
			l_hi3car = 1
		case "No"
			l_hi3car = 2
	end select	
	l_hi3tipdoc = l_rs("hi3tipdoc")
	Select case l_hi3tipdoc
		case "DNI"
			l_hi3tipdoc = 1
		case "LC"
			l_hi3tipdoc = 2
		case "LE"
			l_hi3tipdoc = 3
	end select	
	l_hi3nrodoc = l_rs("hi3nrodoc")
   '************************************************
	l_hi4ape = l_rs("hi4ape")
	l_hi4nom = l_rs("hi4nom")
	l_hi4fecnac = l_rs("hi4fecnac")
	l_hi4viv = l_rs("hi4viv")
	Select case l_hi4viv
		case "Si"
			l_hi4viv = 1
		case "No"
			l_hi4viv = 2
	end select	
	l_hi4car = l_rs("hi4car")
	Select case l_hi4car
		case "Si"
			l_hi4car = 1
		case "No"
			l_hi4car = 2
	end select	
	l_hi4tipdoc = l_rs("hi4tipdoc")
	Select case l_hi4tipdoc
		case "DNI"
			l_hi4tipdoc = 1
		case "LC"
			l_hi4tipdoc = 2
		case "LE"
			l_hi4tipdoc = 3
	end select	
	l_hi4nrodoc = l_rs("hi4nrodoc")
   '************************************************
	l_hi5ape = l_rs("hi5ape")
	l_hi5nom = l_rs("hi5nom")
	l_hi5fecnac = l_rs("hi5fecnac")
	l_hi5viv = l_rs("hi5viv")
	Select case l_hi5viv
		case "Si"
			l_hi5viv = 1
		case "No"
			l_hi5viv = 2
	end select	
	l_hi5car = l_rs("hi5car")
	Select case l_hi5car
		case "Si"
			l_hi5car = 1
		case "No"
			l_hi5car = 2
	end select	
	l_hi5tipdoc = l_rs("hi5tipdoc")
	Select case l_hi5tipdoc
		case "DNI"
			l_hi5tipdoc = 1
		case "LC"
			l_hi5tipdoc = 2
		case "LE"
			l_hi5tipdoc = 3
	end select	
	l_hi5nrodoc = l_rs("hi5nrodoc")
   '************************************************
	l_hi6ape = l_rs("hi6ape")
	l_hi6nom = l_rs("hi6nom") 
	l_hi6fecnac = l_rs("hi6fecnac")
	l_hi6viv = l_rs("hi6viv")
	Select case l_hi6viv
		case "Si"
			l_hi6viv = 1
		case "No"
			l_hi6viv = 2
	end select	
	l_hi6car = l_rs("hi6car")
	Select case l_hi6car
		case "Si"
			l_hi6car = 1
		case "No"
			l_hi6car = 2
	end select	
	l_hi6tipdoc = l_rs("hi6tipdoc")
	Select case l_hi6tipdoc
		case "DNI"
			l_hi6tipdoc = 1
		case "LC"
			l_hi6tipdoc = 2
		case "LE"
			l_hi6tipdoc = 3
	end select	
	l_hi6nrodoc = l_rs("hi6nrodoc")
   '************************************************
	l_priins = l_rs("priins")
	l_prides = l_rs("prides")
	l_prihas = l_rs("prihas")
	l_pritit = l_rs("pritit")
   '************************************************'
	l_secins = l_rs("secins")
	l_secdes = l_rs("secdes")
	l_sechas = l_rs("sechas")
	l_sectit = l_rs("sectit")
   '************************************************
	l_terins = l_rs("terins")
	l_terdes = l_rs("terdes")
	l_terhas = l_rs("terhas")
	l_tertit = l_rs("tertit")
   '************************************************
	l_uniins = l_rs("uniins")
	l_unides = l_rs("unides")
	l_unihas = l_rs("unihas")
	l_unitit = l_rs("unitit")
   '************************************************
	l_posins = l_rs("posins")
	l_posdes = l_rs("posdes")
	l_poshas = l_rs("poshas")
	l_postit = l_rs("postit")
   '************************************************
	l_conins = l_rs("conins")
	l_condes = l_rs("condes")
	l_conhas = l_rs("conhas")
	l_contit = l_rs("contit")
   '************************************************
	l_idinom = l_rs("idinom")
	Select case l_idinom
		case ""
			l_idinom = 0
		case "Inglés"
			l_idinom = 1
		case "Frances"
			l_idinom = 2
		case "Italiano"
			l_idinom = 3
		case "Alemán"
			l_idinom = 4
		case "Portugués"
			l_idinom = 5
	end select
	l_idilee = l_rs("idilee")
	Select case l_idilee
		case ""
			l_idilee = 0
		case "Básico"
			l_idilee = 1
		case "Intermedio"
			l_idilee = 2
		case "Intermedio Avanzado"
			l_idilee = 3
		case "Avanzado"
			l_idilee = 4
		case "Bilingue"
			l_idilee = 5
	end select	
	l_idihab = l_rs("idihab")
	Select case l_idihab
		case ""
			l_idihab = 0
		case "Básico"
			l_idihab = 1
		case "Intermedio"
			l_idihab = 2
		case "Intermedio Avanzado"
			l_idihab = 3
		case "Avanzado"
			l_idihab = 4
		case "Bilingue"
			l_idihab = 5
	end select	
	l_idiesc = l_rs("idiesc")
	Select case l_idiesc
		case ""
			l_idiesc = 0
		case "Básico"
			l_idiesc = 1
		case "Intermedio"
			l_idiesc = 2
		case "Intermedio Avanzado"
			l_idiesc = 3
		case "Avanzado"
			l_idiesc = 4
		case "Bilingue"
			l_idiesc = 5
	end select
   '************************************************
	l_idi2nom = l_rs("idi2nom")
	Select case l_idi2nom
		case ""
			l_idi2nom = 0
		case "Inglés"
			l_idi2nom = 1
		case "Frances"
			l_idi2nom = 2
		case "Italiano"
			l_idi2nom = 3
		case "Alemán"
			l_idi2nom = 4
		case "Portugués"
			l_idi2nom = 5
	end select	
	l_idi2lee = l_rs("idi2lee")
	Select case l_idi2lee
		case ""
			l_idi2lee = 0
		case "Básico"
			l_idi2lee = 1
		case "Intermedio"
			l_idi2lee = 2
		case "Intermedio Avanzado"
			l_idi2lee = 3
		case "Avanzado"
			l_idi2lee = 4
		case "Bilingue"
			l_idi2lee = 5
	end select	
	l_idi2hab = l_rs("idi2hab")
	Select case l_idi2hab
		case ""
			l_idi2hab = 0
		case "Básico"
			l_idi2hab = 1
		case "Intermedio"
			l_idi2hab = 2
		case "Intermedio Avanzado"
			l_idi2hab = 3
		case "Avanzado"
			l_idi2hab = 4
		case "Bilingue"
			l_idi2hab = 5
	end select	
	l_idi2esc = l_rs("idi2esc")
	Select case l_idi2esc
		case ""
			l_idi2esc = 0
		case "Básico"
			l_idi2esc = 1
		case "Intermedio"
			l_idi2esc = 2
		case "Intermedio Avanzado"
			l_idi2esc = 3
		case "Avanzado"
			l_idi2esc = 4
		case "Bilingue"
			l_idi2esc = 5
	end select	
   '************************************************
	l_empant1emp = l_rs("empant1emp")
	l_empant1pue = l_rs("empant1pue")
	l_empant1des = l_rs("empant1des")
	l_empant1has = l_rs("empant1has")
	l_empant1tar = l_rs("empant1tar")
   '************************************************
	l_empant2emp = l_rs("empant2emp")
	l_empant2pue = l_rs("empant2pue")
	l_empant2des = l_rs("empant2des")
	l_empant2has = l_rs("empant2has")
	l_empant2tar = l_rs("empant2tar")
   '************************************************
	l_empant3emp = l_rs("empant3emp")
	l_empant3pue = l_rs("empant3pue")
	l_empant3des = l_rs("empant3des")
	l_empant3has = l_rs("empant3has")
	l_empant3tar = l_rs("empant3tar")
   '************************************************
    l_obs         = l_rs("obs")
'	
'	'response.write "NO es fin de archivo"
'	Response.write "<script>alert('Los datos ya fueron ingresados.');</script>"	
else 
	l_tipo = "A"
	l_ape = ""
	l_nom = ""
	
	l_fecnac = ""
	l_nac = ""
	l_estciv = ""
	l_estcivnom = ""
	l_pue = ""
	l_cal = ""
	l_num = ""
	l_pis = ""
	l_dep = ""
	l_loc = "0"
	l_locdes = ""
	l_pro = ""
	l_prodes = ""
	l_codpos = ""
	l_tel = ""
	l_cel = ""
	l_tipdoc = ""
	l_tipdocdes = ""
	l_nrodoc = ""
	l_cuil = ""
	l_apojub = ""
	l_apojubdes = ""
	l_afjp = ""

	l_padape = ""
	l_padnom = ""
	l_padfecnac = ""
	l_padviv = ""
	l_padcar = ""
	l_padtipdoc = ""
	l_padtipdocdes = ""
	l_padnrodoc = ""

	l_madape = ""
	l_madnom = ""
	l_madfecnac = ""
	l_madviv = ""
	l_madcar = ""
	l_madtipdoc = ""
	l_madtipdocdes = ""
	l_madnrodoc = ""

	l_conape = ""
	l_connom = ""
	l_confecnac = ""
	l_conviv = ""
	l_concar = ""
	l_contipdoc = ""
	l_contipdocdes = ""
	l_connrodoc = ""

	l_hi1ape = ""
	l_hi1nom = ""
	l_hi1fecnac = ""
	l_hi1viv = ""
	l_hi1car = ""
	l_hi1tipdoc = ""
	l_hi1tipdocdes = ""
	l_hi1nrodoc = ""

	l_hi2ape = ""
	l_hi2nom = ""
	l_hi2fecnac = ""
	l_hi2viv = ""
	l_hi2car = ""
	l_hi2tipdoc = ""
	l_hi2tipdocdes = ""
	l_hi2nrodoc = ""

	l_hi3ape = ""
	l_hi3nom = ""
	l_hi3fecnac = ""
	l_hi3viv = ""
	l_hi3car = ""
	l_hi3tipdoc = ""
	l_hi3tipdocdes = ""
	l_hi3nrodoc = ""

	l_hi4ape = ""
	l_hi4nom = ""
	l_hi4fecnac = ""
	l_hi4viv = ""
	l_hi4car = ""
	l_hi4tipdoc = ""
	l_hi4tipdocdes = ""
	l_hi4nrodoc = ""

	l_hi5ape = ""
	l_hi5nom = ""
	l_hi5fecnac = ""
	l_hi5viv = ""
	l_hi5car = ""
	l_hi5tipdoc = ""
	l_hi5tipdocdes = ""
	l_hi5nrodoc = ""

	l_hi6ape = ""
	l_hi6nom = ""
	l_hi6fecnac = ""
	l_hi6viv = ""
	l_hi6car = ""
	l_hi6tipdoc = ""
	l_hi6tipdocdes = ""
	l_hi6nrodoc = ""

	l_priins = ""
	l_prides = ""
	l_prihas = ""
	l_pritit = ""

	l_secins = ""
	l_secdes = ""
	l_sechas = ""
	l_sectit = ""

	l_terins = ""
	l_terdes = ""
	l_terhas = ""
	l_tertit = ""

	l_uniins = ""
	l_unides = ""
	l_unihas = ""
	l_unitit = ""

	l_posins = ""
	l_posdes = ""
	l_poshas = ""
	l_postit = ""

	l_conins = ""
	l_condes = ""
	l_conhas = ""
	l_contit = ""

	l_idinom = ""
	l_idilee = ""
	l_idihab = ""
	l_idiesc = ""

	l_idi2nom = ""
	l_idi2lee = ""
	l_idi2hab = ""
	l_idi2esc = ""

	l_empant1emp = ""
	l_empant1pue = ""
	l_empant1des = ""
	l_empant1has = ""
	l_empant1tar = ""

	l_empant2emp = ""
	l_empant2pue = ""
	l_empant2des = ""
	l_empant2has = ""
	l_empant2tar = ""

	l_empant3emp = ""
	l_empant3pue = ""
	l_empant3des = ""
	l_empant3has = ""
	l_empant3tar = ""
	
	l_obs         = ""

	'response.write "Es fin de archivo"
end if
l_rs.Close
'response.end
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="datos_personal_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<!--
<tr>	 
    <td style="background-color: #FFFFFF;" colspan="6" align="center"><img src="/serviciolocal/shared/images/gen_rep/logomoreno2_tra.jpg" border="0"></td>		
</tr>
<tr>	 
    <td style="background-color: #FFFFFF; font-size: 14pt" colspan="6" align="center">RRHH</td>		
</tr>
<tr>	 
    <td style="background-color: #FFFFFF;" colspan="6" align="center"><b>Bienvenido a la Web de Grupo Glencore. El objetivo de la misma es ....</b></td>		
</tr>
<tr>	 
    <td style="background-color: #FFFFFF;" colspan="6" align="center">&nbsp;</td>		
</tr>
<tr>	 
    <th colspan="6" align="center"><b>CONSIDERACIONES GENERALES</b></th>		
</tr>
<tr>
	<td colspan="4" align="center">
		&nbsp;- Si es la <b>primera vez</b> que ingresa, complete el formulario y presione el Botón de <b>Grabar</b> al pie del mismo.
	</td>
</tr>   
<tr>
	<td colspan="4" align="center">
		&nbsp;- Si Quiere <b>modificar</b> sus Datos ingrese su Nro. de Doc en el cuadro siguiente y presione la Lupa.<br>
 		&nbsp;&nbsp;  Una vez modificados presione el Botón de <b>Grabar</b> ubicado al pie del mismo. &nbsp;&nbsp;<input type="text" name="id" size="10" maxlength="10" value="">&nbsp;<img src="/serviciolocal/shared/images/lupa22.gif" border="0" onclick="Javascript:Validar();" style="cursor: hand;" alt="Obtener Datos del Empleado">				
	</td>
</tr> 
<tr>
	<td colspan="4" align="center">
		&nbsp; Los datos marcados con un <b>*</b> son <b>Obligatorios</b>.						
	</td>
</tr>   
-->

<tr>
	<td colspan="6" style="background-color: #FFFFFF;">
	<table cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td style="background-color: #FFFFFF;"  align="center" width="20%">
			<img src="/serviciolocal/shared/images/gen_rep/logomoreno2_tra.jpg" border="0">
		</td>
		<td style="background-color: #FFFFFF; font-size: 10pt" colspan="6" align="center" width="60%">
			<p style="font-size: 9pt">Bienvenido a la web de </p> 
			<p style="font-size: 16pt">Recursos Humanos</p> 
			
		</td>
		<td style="background-color: #FFFFFF; font-size: 07pt" align="right" width="20%"  >
			<img src="/serviciolocal/shared/images/images44.jpeg" border="0">
		</td>		
	</tr>		
	</table>
	</td>
</tr>

<tr>	 
    <th colspan="6" align="center"><b>CONSIDERACIONES GENERALES</b></th>		
</tr>
<tr>
	<td colspan="6" align="center">
		&nbsp;- Si es la <b>primera vez</b> que ingresa, complete el formulario y presione el Botón de <b>Grabar</b> al pie del mismo.
	</td>
</tr>   
<tr>
	<td colspan="6" align="center">
		&nbsp;- Si Quiere <b>modificar</b> sus Datos ingrese su Nro. de Doc en el cuadro siguiente y presione la Lupa.<br>
 		&nbsp;&nbsp;  Una vez modificados presione el Botón de <b>Grabar</b> ubicado al pie del mismo. &nbsp;&nbsp;<input type="text" name="id" size="10" maxlength="10" value="">&nbsp;<img src="/serviciolocal/shared/images/lupa22.gif" border="0" onclick="Javascript:Validar();" style="cursor: hand;" alt="Obtener Datos del Empleado">				
	</td>
</tr> 
<tr>
	<td colspan="6" align="center">
		&nbsp; Los datos marcados con un <b>*</b> son <b>Obligatorios</b>.						
	</td>
</tr>

<tr>	 
    <th colspan="6" align="center"><b>DATOS PERSONALES</b></th>		
</tr>
<tr>
	<td colspan="6">
	<table>
		<tr>
		    <td height="100%" align="right"><b> * Apellido:</b></td>
			<td height="100%">
				<input type="text" name="ape" size="25" maxlength="25" value="<%= l_ape %>">
			</td>
		    <td height="100%" align="right"><b> * Nombre:</b></td>
			<td height="100%">
				<input type="text" name="nom" size="25" maxlength="25" value="<%= l_nom %>">
			</td>	
		    <td align="right" nowrap width="0"><b> * Fec. Nac.:</b></td>
			<td align="left" nowrap width="0" >
			    <input type="Text" name="fecnac" size="10" maxlength="10" value="<%= l_fecnac %>">
				<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.fecnac)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
			</td>		
		</tr>
		<tr>
			<td height="100%" align="right"><b> * Nacionalidad:</b></td>
			<td height="100%">
				<input type="text" name="nac" size="15" maxlength="25" value="<%= l_nac %>">
			</td>
			<td align="right" nowrap><b> * Est. Civil:</b></td>
			<td>
				<select name="estciv" size="1" style="width:130;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1 >SOLTERO/A</option>
				<option value=2 >CASADO/A</option>
				<option value=3 >DIVORCIADO/A</option>
				<option value=4 >SEPARADO/A</option>
				<option value=5 >VIUDO/A</option>					
				</select>		
				<script> document.datos.estciv.value= "<%= l_estciv %>"</script>									
			</td>
			<td height="100%" align="right"><b> * Puesto:</b></td>
			<td height="100%">
				<input type="text" name="pue" size="15" maxlength="25" value="<%= l_pue %>">
			</td>					
		</tr>		
	</table>
	</td>
</tr>	
<tr>	 
    <th colspan="6" align="center"><b>DOMICILIO</b></th>		
</tr>	
<tr>
	<td colspan="6">
	<table>
	<tr>
	    <td align="right" nowrap width="0"><b> * Calle:</b></td>
		<td align="left" nowrap width="0" >
			<input type="text" name="cal" size="15" maxlength="25" value="<%= l_cal %>">
		</td>
	    <td height="100%" align="right"><b> * Nro:</b></td>
		<td height="100%">
			<input type="text" name="num" size="10" maxlength="10" value="<%= l_num %>">
		</td>
	    <td height="100%" align="right"><b>Piso:</b></td>
		<td height="100%">
			<input type="text" name="pis" size="10" maxlength="10" value="<%= l_pis %>">
		</td>		
	    <td height="100%" align="right"><b>Depto:</b></td>
		<td height="100%">
			<input type="text" name="dep" size="10" maxlength="10" value="<%= l_dep %>">
		</td>		
	</tr>	
	<tr>
	    <td align="right" nowrap><b> * Localidad:</b></td>		
		<td>		
  		   <select name="loc" style="width:170px;" onchange="Javascript:Asignar_Provincia();">
    	   <% Set l_rs = Server.CreateObject("ADODB.RecordSet")%>
	     	       <option value="0" selected>&nbsp;</option>
			   <% l_sql = "SELECT locdes "
				  l_sql = l_sql & " FROM int_localidad"
				  l_sql = l_sql & " ORDER BY locdes"
				  rsOpen l_rs, cn, l_sql, 0 
    			  do while not l_rs.eof
 				  %>
					<option value="<%=l_rs("locdes")%>"><%=l_rs("locdes")%></option>
		 		  <%
					l_rs.MoveNext
				  loop
				  l_rs.close
				  %>
				  <script>document.datos.loc.value = "<%=l_loc%>";</script>
			</select>
		</td>		
		
		<td nowrap align="right" ><b> * Provincia:</b></td>
		<td ta align="left" colspan="3">
			<input type="hidden" name="pro" size="10" maxlength="10" value="<%= l_pro %>">		
		  	<iframe  frameborder="0" name="ifrmpro" scrolling="No" src="provincias_con_00.asp?locdes=<%= l_loc%>&pro=<%= l_pro%>" width="240" height="22"></iframe>
		</td>

	    <td height="100%" align="right"><b> * Cod. Postal:</b></td>
		<td height="100%">
			<input type="text" name="codpos" size="10" maxlength="10" value="<%= l_codpos %>">
		</td>
	    <td colspan="2" >&nbsp;</td>		
	</tr>
	
	<tr>	
		<td height="100%" align="right"><b>Teléfono:</b></td>
		<td height="100%">
			<input type="text" name="tel" size="10" maxlength="20" value="<%= l_tel %>">
		</td>		
	    <td height="100%" align="right"><b>Celular:</b></td>
		<td height="100%">
			<input type="text" name="cel" size="10" maxlength="20" value="<%= l_cel %>">
		</td>		
	    <td colspan="4" >&nbsp;</td>		
	</tr>		
		
	</table>
	</td>							
</tr>	
<tr>	 
    <th colspan="6" align="center"><b>DOCUMENTACION</b></th>		
</tr>
<tr>
	<td colspan="6">
	<table>	   
	<tr>
	    <td align="right" nowrap width="0"><b> * Tipo Documento:</b></td>
		<td align="left" nowrap width="0" >
			<select name="tipdoc" size="1" style="width:80;" >
			<option value=0 selected>&nbsp;</option>
			<option value=1>DNI</option>
			<option value=2>LC</option>
			<option value=3>LE</option>			
			</select>
			<script> document.datos.tipdoc.value= "<%= l_tipdoc %>"</script>						
		</td>
	    <td height="100%" align="right"><b> * Nro:</b></td>
		<td height="100%">
			<input type="text" name="nrodoc" size="10" maxlength="10" value="<%= l_nrodoc %>">
		</td>
	    <td height="100%" align="right"><b>CUIL:</b></td>
		<td height="100%">
			<input type="text" name="cuil" size="15" maxlength="15" value="<%= l_cuil %>">
		</td>		
	</tr>	
	<tr>
	    <td align="right" nowrap width="0"><b> * Aportes Jubilatorios:</b></td>
		<td align="left" nowrap width="0" >
			<select name="apojub" size="1" style="width:120;" >
			<option value=0 selected>&nbsp;</option>
			<option value=1>Reparto</option>
			<option value=2>Capitalización</option>
			</select>	
			<script> document.datos.apojub.value= "<%= l_apojub %>"</script>								
		</td>
	    <td height="100%" align="right"><b>AFJP:</b></td>
		<td height="100%">
			<input type="text" name="afjp" size="25" maxlength="25" value="<%= l_afjp %>">
		</td>
	</tr>	
	</table>
	</td>							
</tr>	
<tr>	 
    <th colspan="6" align="center"><b>DATOS FAMILIARES</b></th>		
</tr>	   
<tr>
	<td colspan="6">
	<table border="0">			
	<tr>	 
	    <td><b>Parentesco</b></td>		
	    <td><b>Apellido</b></td>		
	    <td><b>Nombre</b></td>		
	    <td><b>Fec.Nac.</b></td>		
	    <td><b>Vive</b></td>		
	    <td><b>A Cargo</b></td>		
	    <td><b>Tipo</b></td>												
	    <td><b>Nro. Doc.</b></td>															
	</tr>			
	<tr>
	    <td><b>Padre</b></td>		
	    <td><input type="text" name="padape" size="18" maxlength="25" value="<%= l_padape %>"></td>		
	    <td><input type="text" name="padnom" size="18" maxlength="25" value="<%= l_padnom %>"></td>		
		<td>
		    <input type="Text" name="padfecnac" size="8" maxlength="10" value="<%= l_padfecnac %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.padfecnac)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
	    <td><select name="padviv" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.padviv.value= "<%= l_padviv %>"</script>
		</td>		
	    <td><select name="padcar" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.padcar.value= "<%= l_padcar %>"</script>			
		</td>		
	    <td><select name="padtipdoc" size="1" style="width:50;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>DNI</option>
				<option value=2>LC</option>
				<option value=3>LE</option>			
			</select>
			<script> document.datos.padtipdoc.value= "<%= l_padtipdoc %>"</script>						
		</td>															
	    <td><input type="text" name="padnrodoc" size="7" maxlength="10" value="<%= l_padnrodoc %>"></td>														
	</tr>

	<tr>
	    <td><b>Madre</b></td>		
	    <td><input type="text" name="madape" size="18" maxlength="25" value="<%= l_madape %>"></td>		
	    <td><input type="text" name="madnom" size="18" maxlength="25" value="<%= l_madnom %>"></td>		
		<td>
		    <input type="Text" name="madfecnac" size="8" maxlength="10" value="<%= l_madfecnac %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.madfecnac)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
	    <td><select name="madviv" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.madviv.value= "<%= l_madviv %>"</script>									
		</td>		
	    <td><select name="madcar" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.madcar.value= "<%= l_madcar %>"</script>						
		</td>		
	    <td><select name="madtipdoc" size="1" style="width:50;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>DNI</option>
				<option value=2>LC</option>
				<option value=3>LE</option>			
			</select>
			<script> document.datos.madtipdoc.value= "<%= l_madtipdoc %>"</script>			
		</td>												
	    <td><input type="text" name="madnrodoc" size="7" maxlength="10" value="<%= l_madnrodoc %>"></td>														
	</tr>	
	<tr>
	    <td><b>Conyuge</b></td>		
	    <td><input type="text" name="conape" size="18" maxlength="25" value="<%= l_conape %>"></td>		
	    <td><input type="text" name="connom" size="18" maxlength="25" value="<%= l_connom %>"></td>		
		<td>
		    <input type="Text" name="confecnac" size="8" maxlength="10" value="<%= l_confecnac %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.confecnac)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
	    <td><select name="conviv" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.conviv.value= "<%= l_conviv %>"</script>									
		</td>		
	    <td><select name="concar" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.concar.value= "<%= l_concar %>"</script>						
		</td>		
	    <td><select name="contipdoc" size="1" style="width:50;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>DNI</option>
				<option value=2>LC</option>
				<option value=3>LE</option>			
			</select>
			<script> document.datos.contipdoc.value= "<%= l_contipdoc %>"</script>			
		</td>												
	    <td><input type="text" name="connrodoc" size="7" maxlength="10" value="<%= l_connrodoc %>"></td>														
	</tr>		
	
	<tr>
	    <td><b>Hijo</b></td>		
	    <td><input type="text" name="hi1ape" size="18" maxlength="25" value="<%= l_hi1ape %>"></td>		
	    <td><input type="text" name="hi1nom" size="18" maxlength="25" value="<%= l_hi1nom %>"></td>		
		<td>
		    <input type="Text" name="hi1fecnac" size="8" maxlength="10" value="<%= l_hi1fecnac %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.hi1fecnac)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
	    <td><select name="hi1viv" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi1viv.value= "<%= l_hi1viv %>"</script>												
		</td>		
	    <td><select name="hi1car" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi1car.value= "<%= l_hi1car %>"</script>									
		</td>		
	    <td><select name="hi1tipdoc" size="1" style="width:50;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>DNI</option>
				<option value=2>LC</option>
				<option value=3>LE</option>			
			</select>
			<script> document.datos.hi1tipdoc.value= "<%= l_hi1tipdoc %>"</script>						
		</td>												
	    <td><input type="text" name="hi1nrodoc" size="7" maxlength="10" value="<%= l_hi1nrodoc %>"></td>														
	</tr>	
	
	<tr>
	    <td><b>Hijo</b></td>		
	    <td><input type="text" name="hi2ape" size="18" maxlength="25" value="<%= l_hi2ape %>"></td>		
	    <td><input type="text" name="hi2nom" size="18" maxlength="25" value="<%= l_hi2nom %>"></td>		
		<td>
		    <input type="Text" name="hi2fecnac" size="8" maxlength="10" value="<%= l_hi2fecnac %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.hi2fecnac)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
	    <td><select name="hi2viv" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi2viv.value= "<%= l_hi2viv %>"</script>															
		</td>		
	    <td><select name="hi2car" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi2car.value= "<%= l_hi2car %>"</script>												
		</td>		
	    <td><select name="hi2tipdoc" size="1" style="width:50;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>DNI</option>
				<option value=2>LC</option>
				<option value=3>LE</option>			
			</select>
			<script> document.datos.hi2tipdoc.value= "<%= l_hi2tipdoc %>"</script>									
		</td>												
	    <td><input type="text" name="hi2nrodoc" size="7" maxlength="10" value="<%= l_hi2nrodoc %>"></td>														
	</tr>	
	
	<tr>
	    <td><b>Hijo</b></td>		
	    <td><input type="text" name="hi3ape" size="18" maxlength="25" value="<%= l_hi3ape %>"></td>		
	    <td><input type="text" name="hi3nom" size="18" maxlength="25" value="<%= l_hi3nom %>"></td>		
		<td>
		    <input type="Text" name="hi3fecnac" size="8" maxlength="10" value="<%= l_hi3fecnac %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.hi3fecnac)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
	    <td><select name="hi3viv" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi3viv.value= "<%= l_hi3viv %>"</script>									
		</td>		
	    <td><select name="hi3car" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi3car.value= "<%= l_hi3car %>"</script>						
		</td>		
	    <td><select name="hi3tipdoc" size="1" style="width:50;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>DNI</option>
				<option value=2>LC</option>
				<option value=3>LE</option>			
			</select>
			<script> document.datos.hi3tipdoc.value= "<%= l_hi3tipdoc %>"</script>			
		</td>												
	    <td><input type="text" name="hi3nrodoc" size="7" maxlength="10" value="<%= l_hi3nrodoc %>"></td>														
	</tr>		
		
	<tr>
	    <td><b>Hijo</b></td>		
	    <td><input type="text" name="hi4ape" size="18" maxlength="25" value="<%= l_hi4ape %>"></td>		
	    <td><input type="text" name="hi4nom" size="18" maxlength="25" value="<%= l_hi4nom %>"></td>		
		<td>
		    <input type="Text" name="hi4fecnac" size="8" maxlength="10" value="<%= l_hi4fecnac %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.hi4fecnac)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
	    <td><select name="hi4viv" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi4viv.value= "<%= l_hi4viv %>"</script>									
		</td>		
	    <td><select name="hi4car" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi4car.value= "<%= l_hi4car %>"</script>						
		</td>		
	    <td><select name="hi4tipdoc" size="1" style="width:50;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>DNI</option>
				<option value=2>LC</option>
				<option value=3>LE</option>			
			</select>
			<script> document.datos.hi4tipdoc.value= "<%= l_hi4tipdoc %>"</script>			
		</td>												
	    <td><input type="text" name="hi4nrodoc" size="7" maxlength="10" value="<%= l_hi4nrodoc %>"></td>														
	</tr>		
	
	<tr>
	    <td><b>Hijo</b></td>		
	    <td><input type="text" name="hi5ape" size="18" maxlength="25" value="<%= l_hi5ape %>"></td>		
	    <td><input type="text" name="hi5nom" size="18" maxlength="25" value="<%= l_hi5nom %>"></td>		
		<td>
		    <input type="Text" name="hi5fecnac" size="8" maxlength="10" value="<%= l_hi5fecnac %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.hi5fecnac)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
	    <td><select name="hi5viv" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi5viv.value= "<%= l_hi5viv %>"</script>							
		</td>		
	    <td><select name="hi5car" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi5car.value= "<%= l_hi5car %>"</script>						
		</td>		
	    <td><select name="hi5tipdoc" size="1" style="width:50;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>DNI</option>
				<option value=2>LC</option>
				<option value=3>LE</option>			
			</select>
			<script> document.datos.hi5tipdoc.value= "<%= l_hi5tipdoc %>"</script>			
		</td>												
	    <td><input type="text" name="hi5nrodoc" size="7" maxlength="10" value="<%= l_hi5nrodoc %>"></td>														
	</tr>	
	
	<tr>
	    <td><b>Hijo</b></td>		
	    <td><input type="text" name="hi6ape" size="18" maxlength="25" value="<%= l_hi6ape %>"></td>		
	    <td><input type="text" name="hi6nom" size="18" maxlength="25" value="<%= l_hi6nom %>"></td>		
		<td>
		    <input type="Text" name="hi6fecnac" size="8" maxlength="10" value="<%= l_hi6fecnac %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.hi6fecnac)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
	    <td><select name="hi6viv" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi6viv.value= "<%= l_hi6viv %>"</script>			
		</td>		
	    <td><select name="hi6car" size="1" style="width:45;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Si</option>
				<option value=2>No</option>
			</select>
			<script> document.datos.hi6car.value= "<%= l_hi6car %>"</script>
		</td>		
	    <td><select name="hi6tipdoc" size="1" style="width:50;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>DNI</option>
				<option value=2>LC</option>
				<option value=3>LE</option>			
			</select>
			<script> document.datos.hi6tipdoc.value= "<%= l_hi6tipdoc %>"</script>						
		</td>												
	    <td><input type="text" name="hi6nrodoc" size="7" maxlength="10" value="<%= l_hi6nrodoc %>"></td>														
	</tr>	
	</table>
	</td>							
</tr>	
<tr>	 
    <th colspan="6" align="center"><b>ESTUDIOS FORMALES</b></th>		
</tr>

<tr>
	<td colspan="6">
	<table border="0">			
	<tr>	 
	    <td><b>Tipo</b></td>		
	    <td><b>Institución</b></td>		
	    <td><b>Desde</b></td>		
	    <td><b>Hasta</b></td>		
	    <td><b>Título Obtenido</b></td>		
	</tr>			
	<tr>
	    <td><b>Primario</b></td>		
	    <td><input type="text" name="priins" size="25" maxlength="30" value="<%= l_priins %>"></td>		
		<td>
		    <input type="Text" name="prides" size="10" maxlength="10" value="<%= l_prides %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.prides)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>		
		<td>
		    <input type="Text" name="prihas" size="10" maxlength="10" value="<%= l_prihas %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.prihas)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>				
	    <td><input type="text" name="pritit" size="18" maxlength="25" value="<%= l_pritit %>"></td>				
	</tr>
	<tr>
	    <td><b>Secundario</b></td>		
	    <td><input type="text" name="secins" size="25" maxlength="30" value="<%= l_secins %>"></td>		
		<td>
		    <input type="Text" name="secdes" size="10" maxlength="10" value="<%= l_secdes %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.secdes)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>		
		<td>
		    <input type="Text" name="sechas" size="10" maxlength="10" value="<%= l_sechas %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.sechas)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>				
	    <td><input type="text" name="sectit" size="18" maxlength="25" value="<%= l_sectit %>"></td>				
	</tr>	
	<tr>
	    <td><b>Terciario</b></td>		
	    <td><input type="text" name="terins" size="25" maxlength="30" value="<%= l_terins %>"></td>		
		<td>
		    <input type="Text" name="terdes" size="10" maxlength="10" value="<%= l_terdes %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.terdes)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>		
		<td>
		    <input type="Text" name="terhas" size="10" maxlength="10" value="<%= l_terhas %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.terhas)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>				
	    <td><input type="text" name="tertit" size="18" maxlength="25" value="<%= l_tertit %>"></td>				
	</tr>	
	<tr>
	    <td><b>Universitario</b></td>		
	    <td><input type="text" name="uniins" size="25" maxlength="30" value="<%= l_uniins %>"></td>		
		<td>
		    <input type="Text" name="unides" size="10" maxlength="10" value="<%= l_unides %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.unides)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>		
		<td>
		    <input type="Text" name="unihas" size="10" maxlength="10" value="<%= l_unihas %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.unihas)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>				
	    <td><input type="text" name="unitit" size="18" maxlength="25" value="<%= l_unitit %>"></td>				
	</tr>	
	<tr>
	    <td><b>Postgrado</b></td>		
	    <td><input type="text" name="posins" size="25" maxlength="30" value="<%= l_posins %>"></td>		
		<td>
		    <input type="Text" name="posdes" size="10" maxlength="10" value="<%= l_posdes %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.posdes)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>		
		<td>
		    <input type="Text" name="poshas" size="10" maxlength="10" value="<%= l_poshas %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.poshas)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>				
	    <td><input type="text" name="postit" size="18" maxlength="25" value="<%= l_postit %>"></td>				
	</tr>	
	<tr>
	    <td><b>Otros Conocimientos</b></td>		
	    <td><input type="text" name="conins" size="25" maxlength="30" value="<%= l_conins %>"></td>		
		<td>
		    <input type="Text" name="condes" size="10" maxlength="10" value="<%= l_condes %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.condes)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>		
		<td>
		    <input type="Text" name="conhas" size="10" maxlength="10" value="<%= l_conhas %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.conhas)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>				
	    <td><input type="text" name="contit" size="18" maxlength="25" value="<%= l_contit %>"></td>				
	</tr>	
	</table>
	</td>							
</tr>
<tr>	 
    <th colspan="6" align="center"><b>IDIOMAS</b></th>		
</tr>	   
<tr>
	<td colspan="6">
	<table border="0">			
	<tr>	 
	    <td><b>Lengua</b></td>		
	    <td><b>Lee</b></td>		
	    <td><b>Habla</b></td>		
	    <td><b>Escribe</b></td>		
	</tr>				
	
	<tr>
	    <td><select name="idinom" size="1" style="width:150;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Inglés</option>
				<option value=2>Frances</option>
				<option value=3>Italiano</option>
				<option value=4>Alemán</option>				
				<option value=5>Portugués</option>				
			</select>
			<script> document.datos.idinom.value= "<%= l_idinom %>"</script>			
		</td>		
	    <td><select name="idilee" size="1" style="width:150;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Básico</option>
				<option value=2>Intermedio</option>
				<option value=3>Intermedio Avanzado</option>
				<option value=4>Avanzado</option>				
				<option value=5>Bilingue</option>				
			</select>
			<script> document.datos.idilee.value= "<%= l_idilee %>"</script>						
		</td>		
	    <td><select name="idihab" size="1" style="width:150;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Básico</option>
				<option value=2>Intermedio</option>
				<option value=3>Intermedio Avanzado</option>
				<option value=4>Avanzado</option>				
				<option value=5>Bilingue</option>				
			</select>
			<script> document.datos.idihab.value= "<%= l_idihab %>"</script>			
		</td>		
	    <td><select name="idiesc" size="1" style="width:150;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Básico</option>
				<option value=2>Intermedio</option>
				<option value=3>Intermedio Avanzado</option>
				<option value=4>Avanzado</option>				
				<option value=5>Bilingue</option>				
			</select>
			<script> document.datos.idiesc.value= "<%= l_idiesc %>"</script>			
		</td>		
	</tr>
	<tr>
	    <td><select name="idi2nom" size="1" style="width:150;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Inglés</option>
				<option value=2>Frances</option>
				<option value=3>Italiano</option>
				<option value=4>Alemán</option>				
				<option value=5>Portugués</option>				
			</select>
			<script> document.datos.idi2nom.value= "<%= l_idi2nom %>"</script>			
		</td>		
	    <td><select name="idi2lee" size="1" style="width:150;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Básico</option>
				<option value=2>Intermedio</option>
				<option value=3>Intermedio Avanzado</option>
				<option value=4>Avanzado</option>				
				<option value=5>Bilingue</option>				
			</select>
			<script> document.datos.idi2lee.value= "<%= l_idi2lee %>"</script>			
		</td>		
	    <td><select name="idi2hab" size="1" style="width:150;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Básico</option>
				<option value=2>Intermedio</option>
				<option value=3>Intermedio Avanzado</option>
				<option value=4>Avanzado</option>				
				<option value=5>Bilingue</option>				
			</select>
			<script> document.datos.idi2hab.value= "<%= l_idi2hab %>"</script>			
		</td>		
	    <td><select name="idi2esc" size="1" style="width:150;" >
				<option value=0 selected>&nbsp;</option>
				<option value=1>Básico</option>
				<option value=2>Intermedio</option>
				<option value=3>Intermedio Avanzado</option>
				<option value=4>Avanzado</option>				
				<option value=5>Bilingue</option>				
			</select>
			<script> document.datos.idi2esc.value= "<%= l_idi2esc %>"</script>			
		</td>		
	</tr>	
	</table>
	</td>							
</tr>	
<tr>	 
    <th colspan="6" align="center"><b>EMPLEOS ANTERIORES (INDICAR LOS 3 ULTIMOS EMPLEOS)</b></th>		
</tr>	   
<tr>
	<td colspan="6">
	<table border="0">			
	<tr>	 
	    <td><b>Empresa</b></td>		
	    <td><b>Puesto</b></td>		
	    <td><b>Desde</b></td>		
	    <td><b>Hasta</b></td>		
	    <td><b>Tarea Realizada</b></td>		
	</tr>			
	<tr>
	    <td><input type="text" name="empant1emp" size="20" maxlength="25" value="<%= l_empant1emp %>"></td>		
	    <td><input type="text" name="empant1pue" size="20" maxlength="25" value="<%= l_empant1pue %>"></td>		
		<td>
		    <input type="Text" name="empant1des" size="8" maxlength="10" value="<%= l_empant1des %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.empant1des)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>		
		<td>
		    <input type="Text" name="empant1has" size="8" maxlength="10" value="<%= l_empant1has %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.empant1has)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>				
	    <td><input type="text" name="empant1tar" size="25" maxlength="50" value="<%= l_empant1tar %>"></td>				
	</tr>
	<tr>
	    <td><input type="text" name="empant2emp" size="20" maxlength="25" value="<%= l_empant2emp %>"></td>		
	    <td><input type="text" name="empant2pue" size="20" maxlength="25" value="<%= l_empant2pue %>"></td>		
		<td>
		    <input type="Text" name="empant2des" size="8" maxlength="10" value="<%= l_empant2des %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.empant2des)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>		
		<td>
		    <input type="Text" name="empant2has" size="8" maxlength="10" value="<%= l_empant2has %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.empant2has)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>				
	    <td><input type="text" name="empant2tar" size="25" maxlength="50" value="<%= l_empant2tar %>"></td>				
	</tr>	
	<tr>
	    <td><input type="text" name="empant3emp" size="20" maxlength="25" value="<%= l_empant3emp %>"></td>		
	    <td><input type="text" name="empant3pue" size="20" maxlength="25" value="<%= l_empant3pue %>"></td>		
		<td>
		    <input type="Text" name="empant3des" size="8" maxlength="10" value="<%= l_empant3des %>"> 
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.empant3des)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>		
		<td>
		    <input type="Text" name="empant3has" size="8" maxlength="10" value="<%= l_empant3has %>">
			<a tabindex="-1" href="Javascript:Ayuda_Fecha(document.datos.empant3has)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>				
	    <td><input type="text" name="empant3tar" size="25" maxlength="50" value="<%= l_empant3tar %>"></td>				
	</tr>		
	</table>
	</td>							
</tr>
<tr>	 
    <th colspan="6" align="center"><b>Observaciones u otros datos que considere de Importancia</b></th>		
</tr>
<tr>
	<td colspan="6">
	<table border="0">	   			
	<tr>
	   <td align="center"><textarea name="obs"  rows="3" cols="90" maxlength="500"><%=trim(l_obs)%></textarea>	</td>
	</tr>
	</table>
	</td>							
</tr>

<tr>
    <th colspan="6" align="center" >
		<a tabindex="-1" class=sidebtnABM href="Javascript:Validar_Formulario()">Grabar</a>
		<!--
		<a tabindex="-1" class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
		-->
	</th>
</tr>
<tr>
	<td colspan="6">
		&nbsp;					
	</td>
</tr>   
<tr>
	<td colspan="6">
		&nbsp;					
	</td>
</tr>   
<tr>
	<td colspan="6">
		&nbsp;					
	</td>
</tr>   
</table>
<iframe style="visibility=hidden;" name="valida" src="" width="0%" height="0%"></iframe> 
</form>
<%
set l_rs = nothing
Cn.Close
set Cn = nothing
%>
</body>
</html>
