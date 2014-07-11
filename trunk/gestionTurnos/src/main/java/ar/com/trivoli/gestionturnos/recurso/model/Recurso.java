/**
 * 
 */
package ar.com.trivoli.gestionturnos.recurso.model;

import javax.persistence.Entity;
import javax.persistence.Table;

import ar.com.trivoli.gestionturnos.common.model.EntidadBase;

/**
 * @author posadas
 * 
 *         Entidad que representa un Recurso Reservable del Sistema de Gestion
 *         de Turnos
 */

@Entity
@Table(name = "recursosReservables")
public class Recurso extends EntidadBase<Long> {

	private String descripcion;

	public String getDescripcion() {
		return descripcion;
	}

	public void setDescripcion(String descripcion) {
		this.descripcion = descripcion;
	}
}
