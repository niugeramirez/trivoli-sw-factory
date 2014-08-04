/**
 * 
 */
package ar.com.trivoli.gestionturnos.obrasocial.model;

import ar.com.trivoli.gestionturnos.common.model.EntidadBase;
import javax.persistence.Entity;
import javax.persistence.Table;

/**
 * @author ramirez
 * 
 * 			Entidad que representa las obras sociales del sistema
 *
 */


@Entity
@Table(name = "obrasSociales")
public class ObraSocial extends EntidadBase<Integer> {

	/**
	 * 
	 */
	
	private String nombre;


	public String getNombre() {
		return nombre;
	}

	public void setNombre(String nombre) {
		this.nombre = nombre;
	}

	
}
