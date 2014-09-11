/**
 * 
 */
package ar.com.trivoli.gestionturnos.calendario.model;

import java.util.List;

import ar.com.trivoli.gestionturnos.turno.model.Turno;

/**
 * @author ramirez
 *
 */
public class CalendarioDTO {

	private Calendario calendario;
	private List<Turno> turnos;
	
	
	public Calendario getCalendario() {
		return calendario;
	}
	public void setCalendario(Calendario calendario) {
		this.calendario = calendario;
	}
	public List<Turno> getTurnos() {
		return turnos;
	}
	public void setTurnos(List<Turno> turnos) {
		this.turnos = turnos;
	}
	
	
	
	public CalendarioDTO(Calendario calendario, List<Turno> turnos) {
		super();
		this.calendario = calendario;
		this.turnos = turnos;
	}
		
	
}
