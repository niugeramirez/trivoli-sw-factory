/**
 * 
 */
package ar.com.trivoli.gestionturnos.turno.model;

import javax.persistence.Entity;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.persistence.Table;

import ar.com.trivoli.gestionturnos.calendario.model.Calendario;
import ar.com.trivoli.gestionturnos.common.model.EntidadBase;
import ar.com.trivoli.gestionturnos.paciente.model.Paciente;
import ar.com.trivoli.gestionturnos.practica.model.Practica;

/**
 * @author ramirez
 *
 */

@Entity
@Table(name = "turnos")
public class Turno extends EntidadBase<Integer>{

	@ManyToOne
	@JoinColumn(name = "idCalendario")
	private Calendario 	calendario;

	@ManyToOne
	@JoinColumn(name = "idClientePaciente")
	private Paciente	paciente;

	@ManyToOne
	@JoinColumn(name = "idPractica")
	private Practica	practica;
	
	
	public Calendario getCalendario() {
		return calendario;
	}
	public void setCalendario(Calendario calendario) {
		this.calendario = calendario;
	}
	public Paciente getPaciente() {
		return paciente;
	}
	public void setPaciente(Paciente paciente) {
		this.paciente = paciente;
	}
	public Practica getPractica() {
		return practica;
	}
	public void setPractica(Practica practica) {
		this.practica = practica;
	}
	
	
}
