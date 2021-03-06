/**
 * 
 */
package ar.com.trivoli.gestionturnos.calendario.repository;

import java.util.Date;
import java.util.List;

import org.springframework.data.repository.PagingAndSortingRepository;

import ar.com.trivoli.gestionturnos.calendario.model.Calendario;
import ar.com.trivoli.gestionturnos.recurso.model.Recurso;

/**
 * @author ramirez
 *
 */
public interface ICalendarioRepository extends
		PagingAndSortingRepository<Calendario, Integer>{

	List<Calendario> findByRecurso(Recurso recurso);
	
	//TODO encontrar alguna manera de que el filtro de fecha no se betweeen, sino hacer algo tipo trunc
	List<Calendario> findByRecursoAndFechaHoraInicioBetween(Recurso recurso
															,Date fechaInicial
															,Date fechaFinal);	
	
}
