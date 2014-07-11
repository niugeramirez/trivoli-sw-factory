/**
 * 
 */
package ar.com.trivoli.gestionturnos.modeloturno.repository;

import org.springframework.data.repository.PagingAndSortingRepository;

import ar.com.trivoli.gestionturnos.modeloturno.model.ModeloTurno;

/**
 * @author posadas
 * 
 */
public interface IModeloTurnoRepository extends
		PagingAndSortingRepository<ModeloTurno, Integer> {

}
