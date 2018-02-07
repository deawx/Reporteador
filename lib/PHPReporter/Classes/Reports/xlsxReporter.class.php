<?php

namespace PHPReporter\Reports;

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PHPReporter\Reports\Exception AS Exception;
use PHPReporter\Reports\XLSX\xlsxDrawReport AS DRAW;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

/**
 * XLSX Reporter
 * 
 * Permite generar reportes en formato .xlsx (Excel)
 * 
 * @author Andrés Venegas R.
 * @package xlsxReporter
 * 
 */
class xlsxReporter  {

	/**
     * objExcel.
     *
     * @var Spreadsheet
     */
	private static $objExcel;

	/**
	 * activeSheet
	 * 
	 * @var Active Sheet
	 */
	private static $activeSheet;

	/**
	 * Última fila
	 * 
	 * @var number número de la última fila con datos.
	 * 
	 */
	private static $lastRow;

	/**
	 * Set de datos por rango
	 * 
	 * @var array con las claves data y props ()
	 */
	private static $dataRange;

	/**
	 * Constructor 
	 */
	public function __construct(){
		self::$objExcel = new Spreadsheet();
		self::$activeSheet = self::$objExcel->setActiveSheetIndex(0);
	}

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

		self::setConfigReport($params);

		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="' . $file_name . '.xlsx"');
		header('Cache-Control: max-age=0');
		header('Cache-Control: max-age=1');
		header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
		header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
		header('Cache-Control: cache, must-revalidate');
		header('Pragma: public');

		self::prepareReport($data);

		$writer = IOFactory::createWriter(self::$objExcel, 'Xlsx');
		$writer->save('php://output');
	}

	/**
	 * setConfigReport
	 * 
	 * @param array $data Arreglo con el modelo de datos.
	 * 
	 * @return nothing
	 */
	private function setConfigReport($params){
		self::$objExcel->getProperties()->setCreator($params['autor'])
			->setLastModifiedBy($params['autor'])
			->setTitle($params['titulo'])
			->setSubject($params['asunto'])
			->setDescription($params['descripcion'])
			->setKeywords($params['palabrasclave'])
			->setCategory($params['categoria']);
	}

	/**
	 * prepareRepor: Genera un reporte en formato XLSX (Excel)
	 *
	 * @param array $data Arreglo con el modelo de datos.
	 * 
	 * @return mixed reporte generado
	 * 
	 */
	private function prepareReport($data){
		try {

			if((count($data) < 1) || (gettype($data) !== 'array'))
				throw new Exception("No se obtuvo información para generar el archivo.", 1);

			foreach ($data as $cell => $val) {
				self::$dataRange = $val;
				if(!array_key_exists('range', $val))
					throw new Exception("No se especifico un rango dentro del set de datos.", 1);
				
				if(gettype($val['range']) !== 'string')
					throw new Exception("El rango o celda especificado debe de ser del tipo string.", 1);
					
				self::_checkRange($val['range']);
			}

			
		} catch (Exception $e) {
			self::__errorHandler($e);
		}
	}

	private function _checkRange($range){
		try {
			$draw = new DRAW(self::$dataRange);
			switch (self::__validateCell($range)) {
				case 'FIXED_CELL':
						$draw->FIXED_CELL(self::$activeSheet, $range);
					break;

				case 'VARIABLE_CELL':
						$draw->VARIABLE_CELL(self::$activeSheet, $range);
					break;

				case 'FIXED_RANGE':
						$draw->FIXED_RANGE(self::$activeSheet, $range);
					break;

				case 'VARIABLE_RANGE':
						$draw->VARIABLE_RANGE(self::$activeSheet, $range);
					break;

				case 'FIXED_CELL_TO_N_COL':
						$draw->FIXED_CELL_TO_N_COL(self::$activeSheet, $range);
					break;

				case 'VARIABLE_CELL_TO_N_COL':
						$draw->VARIABLE_CELL_TO_N_COL(self::$activeSheet, $range);
					break;
				
				default:
						throw new Exception("El rango especificado: " . $range . " no es válido.", 1);
					break;
			}
		} catch (Exception $e) {
			self::__errorHandler($e);
		}
	}

	/**
	 * Validar celda: Valida que sea una celda o rango válido
	 * 
	 * @param $cell string celda o rango de celdas
	 * 
	 * @return string con el tipo de celda o rango.
	 * 
	 */
	private function __validateCell($cell){
		if(preg_match("/^[a-zA-Z]+[1-9]+$/", $cell))
			return 'FIXED_CELL';
		else if(preg_match("/^[a-zA-Z]+[$]+$/", $cell))
			return 'VARIABLE_CELL';
		else if(preg_match("/^[a-zA-Z]+[1-9][:][a-zA-Z]+[1-9]+$/", $cell))
			return 'FIXED_RANGE';
		else if(preg_match("/^[a-zA-Z]+[$][:][a-zA-Z]+[$]+$/", $cell))
			return 'VARIABLE_RANGE';
		else if(preg_match("/^[a-zA-Z]+[1-9][:][$]+$/", $cell))
				return 'FIXED_CELL_TO_N_COL';
		else if(preg_match("/^[a-zA-Z]+[$][:][$]+$/", $cell))
				return 'VARIABLE_CELL_TO_N_COL';
		else 
			return '';
	}

	/**
	 * Manejador de errores
	 * 
	 * @return mixed Regresa el archivo con el error generado
	 */
	private function __errorHandler(Exception $e){
		// echo '<br>' . $e->getMessage() . '<br>';
		self::$activeSheet->setCellValue('A1', $e->getMessage());
	}
}