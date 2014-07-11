/**
 * 
 */
package ar.com.trivoli.gestionturnos.recurso.repository;

import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;
import org.springframework.data.repository.PagingAndSortingRepository;

import ar.com.trivoli.gestionturnos.recurso.model.Recurso;

/**
 * @author posadas
 * 
 *         Repositorio de Recursos
 */
public interface IRecursoRepository extends
		PagingAndSortingRepository<Recurso, Integer> {
	Page<Recurso> findByDescripcionLike(Pageable pageable, String descripcion);
}
