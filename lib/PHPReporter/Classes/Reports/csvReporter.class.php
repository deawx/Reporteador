<?php

namespace PHPReporter\Reports;

use PHPReporter\Reports\Exception;

/**
 * CSV Reporter
 * 
 * Permite generar reportes en formato .csv (Texto plano separado por comas).
 * 
 * Ejemplo modelo de datos:
 * 
 *	 $data = array(
 *		'headers' => array('Header 1', 'Header 2', 'Header 1'),
 *		'cols' => array('col_1','col_2','col_3'),
 *		'rows' => array(
 *			array(
 *				'col_1' => value,
 *		        'col_2' => value,
 *		        'col_3' => value
 *			)
 *		)
 *	);
 * 
 * @author Andrés Venegas R.
 * @package csvReporter
 * 
 */
class csvReporter  {

	/**
	 * display: Genera el reporte en formato CSV
	 * 
	 * @param array $params Arreglo con los parámetros del reporte.
	 * @param array $data Arreglo con el modelo de datos del reporte
	 * 
	 * @return file Archivo .csv
	 */
	public static function display($params, $data){
		$file_name = preg_replace(array('/[\s]+/','/[^0-9a-zA-Z\-_\.]/'),array('_',''), $params['nombre']);

		header("Content-type: application/csv");
		header("Content-Disposition: attachment; filename=".$file_name.".csv");
		header("Pragma: no-cache");
		header("Expires: 0");

		echo self::prepareReport($data);
	}

	/**
	 * prepareRepor: Genera la estructura de un reporte csv
	 * acomodando los nombres de las columnas y los renglones
	 * de datos como se necesitan.
	 * 
	 * @param array $data Arreglo con el modelo de datos.
	 * 
	 * @return string Cadena de texto con la información procesada.
	 * 
	 */
	private function prepareReport($data){
		try {
			$headers = $data['headers'];
			$rows = $data['rows'];
			$cols = $data['cols'];
			$arr = array();

			if(count($headers) > count($cols) || count($cols) > count($headers))
				throw new \Exception("Los encabezados no coinciden con el modelo de datos.", 1);

			for($i = 0; $i < count($rows); $i++){
				$arrTmp = array();
				for($j = 0; $j < count($cols); $j++){
					if(array_key_exists($cols[$j], $rows[$i])){
						$arrTmp[$cols[$j]] = $rows[$i][$cols[$j]];
					} else {
						throw new \Exception("La clave " . $cols[$j] . " no se encontro en el modelo de datos.", 1);
					}
				}
				$arr[] = $arrTmp;
			}

			$reporte = "";
			foreach ($headers as $header) {
				$reporte .= ",\"" . $header . "\"";
			}

			$reporte = substr($reporte, 1) . "\n";

			for($i = 0; $i < count($arr); $i++){
				$temp = "";
				foreach ($arr[$i] as $key) {
					$temp .= ",\"" . $key . "\"";
				}
				$reporte .= substr($temp, 1) . "\n";
			}

			return $reporte;	
		} catch (\Exception $e) {
			return $e->getMessage();
		}
	}
}