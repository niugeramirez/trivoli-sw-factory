/**
 * 
 */
package ar.com.trivoli.gestionturnos.calendario.repository;

import org.springframework.data.repository.PagingAndSortingRepository;

import ar.com.trivoli.gestionturnos.calendario.model.Calendario;

/**
 * @author ramirez
 *
 */
public interface ICalendarioRepository extends
		PagingAndSortingRepository<Calendario, Integer>{

}
