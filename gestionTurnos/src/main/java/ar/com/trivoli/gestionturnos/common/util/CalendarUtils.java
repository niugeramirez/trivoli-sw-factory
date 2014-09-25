/**
 * 
 */
package ar.com.trivoli.gestionturnos.common.util;

/**
 * @author ramirez
 *
 */
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Locale;
import java.util.TimeZone;

/**
 * Clase CalendarUtils: Clase Common que brinda utilidades varias para la
 * manipulación de Fechas
 * 
 */
public class CalendarUtils {
	/**
	 * Ejemplo: 201008
	 */
	public static final DateFormat FORMATO_yyyyMM;
	static {
		FORMATO_yyyyMM = new SimpleDateFormat("yyyyMM");
		FORMATO_yyyyMM.setLenient(false);
	}

	/**
	 * Ejemplo: 08/2010
	 */
	public static final DateFormat FORMATO_MMyyyy;
	static {
		FORMATO_MMyyyy = new SimpleDateFormat("MM/yyyy", Locale.ENGLISH);
		FORMATO_MMyyyy.setLenient(false);
	}
	/**
	 * Ejemplo: 20101225
	 */
	public static final DateFormat FORMATO_yyyyMMdd;
	static {
		FORMATO_yyyyMMdd = new SimpleDateFormat("yyyyMMdd");
		FORMATO_yyyyMMdd.setLenient(false);
	}

	/**
	 * Ejemplo: 2010-12-25
	 */
	public static final DateFormat FORMATO_yyyyMMdd_ConSeparador;
	static {
		FORMATO_yyyyMMdd_ConSeparador = new SimpleDateFormat("yyyy-MM-dd");
		FORMATO_yyyyMMdd_ConSeparador.setLenient(false);
	}

	/**
	 * Ejemplo: 2010/12/25 00:00:00
	 */
	public static final DateFormat FORMATO_yyyyMMdd_hhmmss;
	static {
		FORMATO_yyyyMMdd_hhmmss = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		FORMATO_yyyyMMdd_hhmmss.setLenient(false);
	}

	/**
	 * Ejemplo: 25/12/2010
	 */
	public static final DateFormat FORMATO_ddMMyyyy;
	static {
		FORMATO_ddMMyyyy = new SimpleDateFormat("dd/MM/yyyy");
		FORMATO_ddMMyyyy.setLenient(false);
	}

	/**
	 * Ejemplo: 25/12
	 */
	public static final DateFormat FORMATO_ddMM;
	static {
		FORMATO_ddMM = new SimpleDateFormat("dd/MM");
		FORMATO_ddMM.setLenient(false);
	}

	/**
	 * Ejemplo: 25/12/2010 00:00:00
	 */
	public static final DateFormat FORMATO_ddMMyyyy_hhmmss;
	static {
		FORMATO_ddMMyyyy_hhmmss = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
		FORMATO_ddMMyyyy_hhmmss.setLenient(false);
	}

	/**
	 * Ejemplo: 25 de Diciembre de 2010
	 */
	public static final DateFormat FORMATO_DESCRIPCION_SP;
	static {
		FORMATO_DESCRIPCION_SP = new SimpleDateFormat("dd 'de' MMMM 'de' yyyy",
				new Locale("ES"));
		FORMATO_DESCRIPCION_SP.setLenient(false);
	}

	/**
	 * Ejemplo: 10
	 */
	public static final DateFormat FORMATO_yy;
	static {
		FORMATO_yy = new SimpleDateFormat("yy");
		FORMATO_yy.setLenient(false);
	}

	/**
	 * Dada una Fecha, lleva sus Horas, Minutos y Segundos al valor máximo
	 * (23:59:59)
	 * 
	 * @param fecha
	 *            Parametro Fecha
	 * @return Fecha resultante con sus Horas, Minutos y Segundos modificados
	 */
	public static Date getFechaLimiteSuperior(Date fecha) {
		Calendar calendar = GregorianCalendar.getInstance(TimeZone
				.getTimeZone("GMT-03:00"));

		calendar.setTime(fecha);
		calendar.add(GregorianCalendar.DATE, 1);

		calendar.set(GregorianCalendar.AM_PM, GregorianCalendar.AM);

		calendar.set(GregorianCalendar.HOUR, 0);
		calendar.set(GregorianCalendar.MINUTE, 0);
		calendar.set(GregorianCalendar.SECOND, 0);

		calendar.add(GregorianCalendar.SECOND, -1);

		return calendar.getTime();
	}

	/**
	 * Dada una Fecha, lleva sus Horas, Minutos, Segundos y Milisegundos al
	 * valor mínimo (00:00:00.0000)
	 * 
	 * @param fecha
	 *            Parametro Fecha
	 * @return Fecha resultante con sus Horas, Minutos, Segundos y Milisegundos
	 *         modificados
	 */
	public static Date getFechaLimiteInferior(Date fecha) {
		Calendar calendar = GregorianCalendar.getInstance(TimeZone
				.getTimeZone("GMT-03:00"));

		calendar.setTime(fecha);

		calendar.set(GregorianCalendar.AM_PM, GregorianCalendar.AM);
		calendar.set(GregorianCalendar.HOUR, 0);
		calendar.set(GregorianCalendar.MINUTE, 0);
		calendar.set(GregorianCalendar.SECOND, 0);
		calendar.set(GregorianCalendar.MILLISECOND, 0);

		return calendar.getTime();
	}

	/**
	 * Dada un String que representa una Fecha un un dado Formato, devuelve la
	 * Fecha representada, convertida a un String en otro Formato
	 * 
	 * @param input
	 *            String que representa una Fecha
	 * @param inputFormat
	 *            Formato de Entrada
	 * @param outputFormat
	 *            Formato de Salida
	 * @return
	 */
	public static String parseToFormat(String input, DateFormat inputFormat,
			DateFormat outputFormat) {
		String result = null;
		try {
			result = outputFormat.format(inputFormat.parse(input));
		} catch (ParseException e) {
			throw new RuntimeException("Error en parseo de fecha", e);
		}
		return result;
	}

	/**
	 * Dada una una Fecha devuelve la misma convertida a un String en un Formato
	 * dado
	 * 
	 * @param input
	 *            Date que representa una Fecha
	 * @param outputFormat
	 *            Formato de Salida
	 * @return
	 */
	public static String parseToFormat(Date input, DateFormat outputFormat) {
		String result = null;
		result = outputFormat.format(input);
		return result;
	}

	/**
	 * Dado un Periodo en Formato YYYYMM (Ej:201010), calcula el Periodo
	 * inmediato anterior en el mismo Formato (Ej:201009)
	 * 
	 * @param periodo
	 *            Periodo
	 * @return Periodo Previo
	 */
	public static Integer calcularPeriodoPrevio(Integer periodo) {
		Calendar calendar = GregorianCalendar.getInstance(TimeZone
				.getTimeZone("GMT-03:00"));

		Integer periodoPrevio = null;

		try {
			calendar.setTime(CalendarUtils.FORMATO_yyyyMM.parse(periodo
					.toString()));
			calendar.add(Calendar.MONTH, -1);

			periodoPrevio = Integer.parseInt(CalendarUtils.parseToFormat(
					calendar.getTime(), CalendarUtils.FORMATO_yyyyMM));

		} catch (ParseException e) {
			throw new RuntimeException("Error de Parseo de Periodo: "
					+ periodo.toString().trim() + ".");
		}
		return periodoPrevio;
	}

	/**
	 * Calcula la diferencia en tiempo entre dos fechas, dependiendo del valor
	 * por parametro {@code milisegundosDeTruncado}. En caso de ser nulo, se
	 * toma 1 (diferencia en milisegundos).
	 * 
	 * @param fecha1
	 *            la fecha1
	 * @param fecha2
	 *            la fecha2
	 * @param milisegundosDeTruncado
	 *            los milisegundos para el truncado
	 * @return the long
	 */
	public static long calcularDiferencia(Date fecha1, Date fecha2,
			Long milisegundosDeTruncado) {
		Calendar cal = Calendar.getInstance();
		cal.setTime(fecha1);
		cal.set(Calendar.HOUR_OF_DAY, 0);
		cal.set(Calendar.MINUTE, 0);
		cal.set(Calendar.MILLISECOND, 0);
		long tmFechaVencimiento = cal.getTimeInMillis();
		cal.setTime(fecha2);
		cal.set(Calendar.HOUR_OF_DAY, 0);
		cal.set(Calendar.MINUTE, 0);
		cal.set(Calendar.MILLISECOND, 0);
		long tmFechaPronostico = cal.getTimeInMillis();
		long diff = tmFechaVencimiento - tmFechaPronostico;
		int daysDiff = (int) (diff / (milisegundosDeTruncado != null ? milisegundosDeTruncado
				: 1));
		return daysDiff;
	}

	public static boolean isWeekend(Date fechaAGenerar) {
		Calendar cal = Calendar.getInstance();
		cal.setTime(fechaAGenerar);

		if (cal.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY
				|| cal.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY) {
			return true;
		}

		return false;
	}

}
