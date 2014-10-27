/**
 * 
 */
package ar.com.trivoli.gestionturnos.paciente.model;

import javax.persistence.Entity;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.persistence.Table;

import ar.com.trivoli.gestionturnos.common.model.EntidadBase;
import ar.com.trivoli.gestionturnos.obrasocial.model.ObraSocial;




/**
 * @author ramirez
 *
 *	Entidad que representa los Clientes/Pacientes del sistema
 *
 *
 */
@Entity
@Table(name = "clientesPacientes")
public class Paciente extends EntidadBase<Integer> {

	private String dni;
	private	String nombre;
	private	String apellido;
	private	String nroHistoriaClinica;

	@ManyToOne
	@JoinColumn(name = "idObraSocial")
	private ObraSocial obraSocial;
	
	public ObraSocial getObraSocial() {
		return obraSocial;
	}

	public void setObraSocial(ObraSocial obraSocial) {
		this.obraSocial = obraSocial;
	}

	public String getDni() {
		return dni;
	}
	
	public void setDni(String dni) {
		this.dni = dni;
	}
	
	public String getNombre() {
		return nombre;
	}
	
	public void setNombre(String nombre) {
		this.nombre = nombre;
	}
	
	public String getApellido() {
		return apellido;
	}
	
	public void setApellido(String apellido) {
		this.apellido = apellido;
	}
	
	public String getNroHistoriaClinica() {
		return nroHistoriaClinica;
	}
	
	public void setNroHistoriaClinica(String nroHistoriaClinica) {
		this.nroHistoriaClinica = nroHistoriaClinica;
	}
	
	
}
