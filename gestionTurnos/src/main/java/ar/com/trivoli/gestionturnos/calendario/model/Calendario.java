/**
 * 
 */
package ar.com.trivoli.gestionturnos.calendario.model;

import java.util.Date;

import javax.persistence.Entity;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.persistence.Table;

import ar.com.trivoli.gestionturnos.common.model.EntidadBase;
import ar.com.trivoli.gestionturnos.recurso.model.Recurso;



/**
 * @author ramirez
 *
 *	Calendario de turnos libres y otorgados
 *
 *
 */

@Entity
@Table(name = "calendario")
public class Calendario extends EntidadBase<Integer> {
	
	private Date fechaHoraInicio;
	private Date fechaHoraFin;
	
	@ManyToOne
	@JoinColumn(name = "idRecursoReservable")
	private Recurso recurso;

	
	
	public Date getFechaHoraInicio() {
		return fechaHoraInicio;
	}

	public void setFechaHoraInicio(Date fechaHoraInicio) {
		this.fechaHoraInicio = fechaHoraInicio;
	}

	public Date getFechaHoraFin() {
		return fechaHoraFin;
	}

	public void setFechaHoraFin(Date fechaHoraFin) {
		this.fechaHoraFin = fechaHoraFin;
	}

	public Recurso getRecurso() {
		return recurso;
	}

	public void setRecurso(Recurso recurso) {
		this.recurso = recurso;
	}

	
	
}
