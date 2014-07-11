/**
 * 
 */
package ar.com.trivoli.gestionturnos.common.model;

import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.MappedSuperclass;

/**
 * @author posadas
 * 
 *         Entidad Base de la Aplicación parametrizada por el Tipo de ID
 * 
 */
@MappedSuperclass
public abstract class EntidadBase<tipoId> {

	@Id
	@GeneratedValue
	protected tipoId id;

	public tipoId getId() {
		return id;
	}
}
