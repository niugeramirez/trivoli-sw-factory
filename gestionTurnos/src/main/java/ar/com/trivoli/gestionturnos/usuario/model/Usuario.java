/**
 * 
 */
package ar.com.trivoli.gestionturnos.usuario.model;

import javax.persistence.Entity;
import javax.persistence.Table;

import ar.com.trivoli.gestionturnos.common.model.EntidadBase;

/**
 * @author posadas
 * 
 *         Tabla de Usuarios de la Aplicacion
 * 
 */
@Entity
@Table(name = "usuarios")
public class Usuario extends EntidadBase<Integer> {
	private String usuario;

	private String nombre;

	private String apellido;

	public String getApellido() {
		return apellido;
	}

	public String getNombre() {
		return nombre;
	}

	public String getNombreCompleto() {
		return this.apellido + ", " + this.nombre;
	}

	public String getUsuario() {
		return usuario;
	}

	public void setApellido(String apellido) {
		this.apellido = apellido;
	}

	public void setNombre(String nombre) {
		this.nombre = nombre;
	}

	public void setUsuario(String usuario) {
		this.usuario = usuario;
	}

}
