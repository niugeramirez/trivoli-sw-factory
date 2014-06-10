/**
 * 
 */
package ar.com.trivoli.gestionturnos.recurso.model;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.Table;

/**
 * @author posadas
 * 
 *         Entidad que representa un Recurso Reservable del Sistema de Gestion
 *         de Turnos
 */

@Entity
@Table(name = "recursosReservables")
public class Recurso {

	@Id
	@GeneratedValue
	private int id;

	private String descripcion;

	public String getDescripcion() {
		return descripcion;
	}

	public int getId() {
		return id;
	}

	public void setDescripcion(String descripcion) {
		this.descripcion = descripcion;
	}

	public void setId(int id) {
		this.id = id;
	}

}
