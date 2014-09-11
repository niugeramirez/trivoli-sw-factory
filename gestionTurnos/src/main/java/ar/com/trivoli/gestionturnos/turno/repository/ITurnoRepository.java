/**
 * 
 */
package ar.com.trivoli.gestionturnos.turno.repository;

import java.util.List;

import org.springframework.data.repository.PagingAndSortingRepository;

import ar.com.trivoli.gestionturnos.calendario.model.Calendario;
import ar.com.trivoli.gestionturnos.turno.model.Turno;

/**
 * @author ramirez
 *
 */
public interface ITurnoRepository extends
					PagingAndSortingRepository<Turno, Integer>{
	List<Turno> findByCalendario(Calendario calendario);
}
