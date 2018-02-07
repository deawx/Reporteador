<?php

namespace PHPReporter;

require_once __DIR__ . '/vendor/autoload.php';

use PHPReporter\Reports\Exception;

class PHPReporter{

	private static $params_default = array(
		'nombre' => 'Ejemplo',
		'titulo' => 'Titulo de ejemplo',
		'autor' => 'Desconocido',
		'asunto' => '',
		'descripcion' => '',
		'palabrasclave' => '',
		'categoria' => ''
	);

	public static function displayReport($type = null, $data = array(), $params = array()){
		try {

			if(is_null($type)) throw new Exception("No se ha especificado el tipo a exportar", 1);

			$params = array_merge(self::$params_default, $params);

			if(in_array($type, array('csv', 'xlsx', 'pdf'))){
				$classname = 'PHPReporter\Reports\\' . $type . "Reporter";
				$report = new $classname();
				$report->display($params, $data);
			} else throw new \Exception("El tipo de reporte especificado no esta disponible.", 1);

		} catch (Exception $e) {
			echo $e->getMessage(), "\n";
		}
	}
}