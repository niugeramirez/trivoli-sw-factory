/**
 * 
 */
package ar.com.trivoli.gestionturnos.practica.model;

import javax.persistence.Entity;
import javax.persistence.Table;

import ar.com.trivoli.gestionturnos.common.model.EntidadBase;

/**
 * @author ramirez
 *
 */

@Entity
@Table(name = "practicas")
public class Practica extends EntidadBase<Integer>{
	
	private String titulo;

	
	public String getTitulo() {
		return titulo;
	}

	public void setTitulo(String titulo) {
		this.titulo = titulo;
	}

}
