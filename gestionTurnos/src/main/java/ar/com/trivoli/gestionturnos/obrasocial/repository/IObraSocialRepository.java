/**
 * 
 */
package ar.com.trivoli.gestionturnos.obrasocial.repository;

import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;
import org.springframework.data.repository.PagingAndSortingRepository;

import ar.com.trivoli.gestionturnos.obrasocial.model.ObraSocial;


/**
 * @author ramirez
 *		
 *			Repositorio de Obras Sociales
 */
public interface IObraSocialRepository	extends
	PagingAndSortingRepository<ObraSocial, Integer>	{
	Page<ObraSocial> findByNombreLike(Pageable pageable, String nombre);
}
