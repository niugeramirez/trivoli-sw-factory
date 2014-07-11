/**
 * 
 */
package ar.com.trivoli.gestionturnos.modeloturno.model;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Table;

import ar.com.trivoli.gestionturnos.common.model.EntidadBase;

/**
 * @author posadas
 * 
 *         Entidad que representa un Modelo de Turnos (Template) del Sistema de
 *         Gestion de Turnos
 */
@Entity
@Table(name = "templateReservas")
public class ModeloTurno extends EntidadBase<Long> {
	private String descripcion;

	@Column(name = "titulo")
	private String detalle;

	public String getDescripcion() {
		return descripcion;
	}

	public String getDetalle() {
		return detalle;
	}

	public void setDescripcion(String descripcion) {
		this.descripcion = descripcion;
	}

	public void setDetalle(String detalle) {
		this.detalle = detalle;
	}
}
