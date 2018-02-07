<?php

namespace PHPReporter\Reports\XLSX;

use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PHPReporter\Reports\Exception AS Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;

/**
 * XLSX Reporter
 * 
 * @author Andrés Venegas R.
 * @package xlsxReporter
 * 
 */
class xlsxDrawReport extends \PHPReporter\Reports\XLSX\xlsxProps {

	private static $data;
	private static $props;
	private static $props_text;
	private static $props_cell;
	private static $colors = array('BLACK','WHITE','RED','DARKRED','BLUE','DARKBLUE','GREEN','DARKGREEN','YELLOW','DARKYELLOW');
	private static $abc = array();

	public function __construct($data){
		if(!array_key_exists("data", $data))
			throw new Exception("Ocurrió un error en la ejecución.", 1);
		self::$data = $data['data'];
		self::$props = array_merge(xlsxProps::$propiedades, (array_key_exists("props", $data)) ? $data['props'] : array());

		if(array_key_exists("props", $data)){
			self::$props_text = array_merge(xlsxProps::$propiedades['text'], ((array_key_exists("text", $data['props'])) ? $data['props']['text'] : array()));
			self::$props_cell = array_merge(xlsxProps::$propiedades['cell'], ((array_key_exists("cell", $data['props'])) ? $data['props']['cell'] : array()));
		} else {
			self::$props_text = array_merge(xlsxProps::$propiedades['text'], array());
			self::$props_cell = array_merge(xlsxProps::$propiedades['cell'], array());	
		}
		self::$abc = array();
		for($i=65; $i<=90; $i++) {  
		    self::$abc[] = chr($i);
		}
	}

	public static function FIXED_CELL(&$sheet, $cell){
		$rt = self::setPropsText(self::$data);
		$sheet->setCellValue($cell, $rt);
		self::setPropsCell($sheet, $cell);
	}

	public static function VARIABLE_CELL(&$sheet, $cell){
		$row = $sheet->getHighestRow() + 1;
		$celda = explode('$', $cell)[0] . $row;
		$rt = self::setPropsText(self::$data);
		$sheet->setCellValue($celda, $rt);
	}

	public static function FIXED_RANGE(&$sheet, $range){
		if(self::$props_cell['merge']){
			self::setPropsCell($sheet, $range, true);
			$cells = explode(':', $range);
			$rt = self::setPropsText(self::$data);
			$sheet->setCellValue($cells[0], $rt);
		} else {
			if(gettype(self::$data) !== 'array')
				throw new Exception('Se necesita un set de datos para el rango: ' . $range, 1);

			$celdas = explode(':', $range);
			$row = self::__prepareRangeCell($celdas[0], true);
			$cell1 = self::__prepareRangeCell($celdas[0]);
			$cell2 = self::__prepareRangeCell($celdas[1]);

			$start = ((count($cell1)-1) * 26) + (array_search($cell1[count($cell1)-1], self::$abc));
			$end = ((count($cell2)-1) * 26) + (array_search($cell2[count($cell2)-1], self::$abc));

			for($i = $start; $i <= $end; $i++){
				$rt = self::setPropsText(self::$data[$i]);
				$sheet->setCellValue(self::$abc[$i].$row, $rt);
				self::setPropsCell($sheet, self::$abc[$i].$row);
			}
		}
	}

	public static function VARIABLE_RANGE(&$sheet, $range){
		if(self::$props_cell['merge']){
			$row = $sheet->getHighestRow() + 1;
			$rango = str_replace("$", $row, $range);
			self::setPropsCell($sheet, $rango, true);
			$cells = explode(':', $rango);
			$rt = self::setPropsText(self::$data);
			$sheet->setCellValue($cells[0], $rt);
		} else {
			$row = $sheet->getHighestRow() + 1;
			$rango = str_replace("$", $row, $range);

			if(gettype(self::$data) !== 'array')
				throw new Exception('Se necesita un set de datos para el rango: ' . $rango, 1);

			$celdas = explode(':', $rango);
			$row = self::__prepareRangeCell($celdas[0], true);
			$cell1 = self::__prepareRangeCell($celdas[0]);
			$cell2 = self::__prepareRangeCell($celdas[1]);

			$start = ((count($cell1)-1) * 26) + (array_search($cell1[count($cell1)-1], self::$abc));
			$end = ((count($cell2)-1) * 26) + (array_search($cell2[count($cell2)-1], self::$abc));

			for($i = $start; $i <= $end; $i++){
				$rt = self::setPropsText(self::$data[$i]);
				$sheet->setCellValue(self::$abc[$i].$row, $rt);
				self::setPropsCell($sheet, self::$abc[$i].$row);
			}
		}
	}

	public static function FIXED_CELL_TO_N_COL(&$sheet, $range){
		$celda = explode(':', $range)[0];
		$nCells = count(self::$data);
		$col = implode(self::__prepareRangeCell($celda));
		$cellStart = self::__prepareRangeCell($celda, true);

		$cell1 = self::__prepareRangeCell($celda);
		$start = ((count($cell1)-1) * 26) + (array_search($cell1[count($cell1)-1], self::$abc));

		for($i = $start; $i < ($start + $nCells); $i++){
			for($j = $start; $j <= count(self::$data[$i - $start]) - 1; $j++){
				$rt = self::setPropsText(floatval(self::$data[$i - $start][$j]));
				$sheet->setCellValue(self::$abc[$j].($i+$cellStart), $rt);
				self::setPropsCell($sheet, self::$abc[$j].($i+$cellStart));
			}
		}
	}

	public static function VARIABLE_CELL_TO_N_COL(&$sheet, $range){
		$celda = explode(':', $range)[0];
		$nCells = count(self::$data);
		$col = implode(self::__prepareRangeCell($celda));
		$cellStart = $sheet->getHighestRow() + 1;

		$cell1 = self::__prepareRangeCell($celda);
		$start = ((count($cell1)-1) * 26) + (array_search($cell1[count($cell1)-1], self::$abc));

		for($i = $start; $i < ($start + $nCells); $i++){
			for($j = $start; $j <= count(self::$data[$i - $start]) - 1; $j++){

				$rt = self::setPropsText(floatval(self::$data[$i - $start][$j]));
				$sheet->setCellValue(self::$abc[$j].($i+$cellStart), $rt);
				self::setPropsCell($sheet, self::$abc[$j].($i+$cellStart));
			}
		}
	}

	private function setPropsCell(&$sheet, $cell, $isRange = false){
		if(!$isRange){
			$sheet->getStyle($cell)->getAlignment()->setWrapText(self::$props_cell['wrapText']);
			$cells = preg_split('//', $cell, -1, PREG_SPLIT_NO_EMPTY);
			$nCell = '';
			foreach ($cells as $let) {
				if(preg_match("/^[a-zA-Z]+$/", $let)){
					$nCell .= $let;
				}
			}
			$sheet->getColumnDimension($nCell)->setAutoSize(true);
		} else {
			if(self::$props_cell['merge']){
				$sheet->getStyle(explode(':', $cell)[0])->getAlignment()->setWrapText(self::$props_cell['wrapText']);
				$sheet->mergeCells($cell);
			}
		}
	}

	private function setPropsText($text){
		try {
			$rt = new RichText();
			$txt = $rt->createTextRun($text);
			if(self::$props_text['color'] !== ""){
				if(in_array(strtoupper(self::$props_text['color']), self::$colors)){
					self::__setFontColor($txt, ('COLOR_' . strtoupper(self::$props_text['color'])));
				}
			} else {
				if(self::$props_text['colorHex'] !== "")
					$txt->getFont()->setColor(new Color(self::$props_text['colorHex']));
			}
			if(self::$props_text['bold'] || !self::$props_text['bold'])
				$txt->getFont()->setBold(((self::$props_text['bold']) ? true : false));
			if(self::$props_text['italic'] || !self::$props_text['italic'])
				$txt->getFont()->setItalic(((self::$props_text['italic']) ? true : false));
			if(self::$props_text['underlined'] || !self::$props_text['underlined'])
				$txt->getFont()->setUnderline(((self::$props_text['underlined']) ? true : false));
			$txt->getFont()->setName((self::$props_text['font'] === '') ? xlsxProps::$propiedades['text']['font'] : self::$props_text['font']);
			$txt->getFont()->setSize((self::$props_text['size'] === '') ? xlsxProps::$propiedades['text']['size'] : self::$props_text['size']);
			return $rt;
		} catch (Exception $e) {
			return $e;
		}
	}

	private function __setFontColor(&$text, $color){
		try {
			switch ($color) {
				case 'COLOR_BLACK':
						$text->getFont()->setColor(new Color(Color::COLOR_BLACK));
					break;
				case 'COLOR_WHITE':
						$text->getFont()->setColor(new Color(Color::COLOR_WHITE));
					break;
				case 'COLOR_RED':
						$text->getFont()->setColor(new Color(Color::COLOR_RED));
					break;
				case 'COLOR_DARKRED':
						$text->getFont()->setColor(new Color(Color::COLOR_DARKRED));
					break;
				case 'COLOR_BLUE':
						$text->getFont()->setColor(new Color(Color::COLOR_BLUE));
					break;
				case 'COLOR_DARKBLUE':
						$text->getFont()->setColor(new Color(Color::COLOR_DARKBLUE));
					break;
				case 'COLOR_GREEN':
						$text->getFont()->setColor(new Color(Color::COLOR_GREEN));
					break;
				case 'COLOR_DARKGREEN':
						$text->getFont()->setColor(new Color(Color::COLOR_DARKGREEN));
					break;
				case 'COLOR_YELLOW':
						$text->getFont()->setColor(new Color(Color::COLOR_YELLOW));
					break;
				case 'COLOR_DARKYELLOW':
						$text->getFont()->setColor(new Color(Color::COLOR_DARKYELLOW));
					break;

				default:
						$text->getFont()->setColor(Color::COLOR_BLACK);
					break;
			}
		} catch (Exception $e) {
			return $e;
		}
	}

	private function __prepareRangeCell($cell, $isRow = false){
		if(!$isRow){
			$cells = preg_split('//', $cell, -1, PREG_SPLIT_NO_EMPTY);
			$range = array();
			foreach ($cells as $let) {
				if(preg_match("/^[a-zA-Z]+$/", $let)){
					array_push($range, $let);
				}
			}
			return $range;
		} else {
			$cells = preg_split('//', $cell, -1, PREG_SPLIT_NO_EMPTY);
			$row = '';
			foreach ($cells as $let) {
				if(preg_match("/^[1-9]+$/", $let)){
					$row .= $let;
				}
			}
			return $row;
		}
	}
}