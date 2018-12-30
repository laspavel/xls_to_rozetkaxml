<?php

define('INPUTFILE','pricelist.xlsx');
define('OUTPUTFILE','pricelist.xml');

define("SHOP_NAME", "Магазин 22 Градуса");
define("SHOP_DESCRIPTION", "Фізична особа-підприємець Теплінська Катерина Андріївна. Интернет-магазин 22 Градуса – официальный дилер завода производителя ТМ UDEN-S");
define("SHOP_URL", "https://shop.uden-s.ua/");

//===========================

require_once 'vendor/PHPExcel.php';

class XMLtoYaml {

	private $products=array();
	private $categories=array();

	public function __construct($inputfile)  
	// При старте считываем лист с товарами и категориями

	{
		$xlsData = $this->getXLS($inputfile);
		foreach ($xlsData[1] as $key => $cat) {
			if ($key==0) continue;
			if ($cat[0]!='') {
				$this->categories[$cat[0]]=array(
					'name'=>trim($cat[1]),
					'parent_id'=>$cat[3],
				);
			}
		}

		$this->products=$xlsData[0];
		unset($xlsData);
	}

	private function getCategoryId($name) 
	// Получение ID категории по имени
	{

		foreach ($this->categories as $catId => $category) {
			if ($category['name']==trim($name)) return $catId;
		}

		die(sprintf("ВНИМАНИЕ !!! КАТЕГОРИЯ %s УКАЗАННАЯ В ТОВАРЕ НЕ НАЙДЕНА В СПИСКЕ КАТЕГОРИЙ !!!\n",$name));
	}

	private function getProductParams($heads,$product) 
	// Получение атрибутов товаров
	
	{
		$params=array();

		foreach ($heads as $key => $head) {
			$paramname=@explode('[', explode(']', $head)[0])[1];
			if (!is_null($paramname)) {
				if (trim($product[$key])!='') {
					$params[$paramname]=trim($product[$key]);
				}
			}
		}

		return $params;

	}

	public function getXMLData() {
	// Получаем данные для обработки		

		$xml_data=array();

		foreach ($this->products as $key => $product) {
				if ($key==0) continue;
				$xml_data[$product[0]]=array(
					'url'=>htmlspecialchars($product[13]),
					'price'=>$product[4],
					'currencyId'=>$product[5],
					'categoryId'=>$this->getCategoryId($product[9]),
					'pictures'=>explode(',',$product[7]),
					'vendor'=>htmlspecialchars($product[11]),
					'stock_quantity'=>$product[8],
					'name'=>htmlspecialchars($product[2]),
					'description'=>"<![CDATA[".nl2br(htmlspecialchars($product[3]))."]]>",
					'params'=>$this->getProductParams($this->products[0],$product),
				);
		}
		
		return $this->SaveXML(OUTPUTFILE,$xml_data);
	}


	private function SaveXML($file,$outdata) 
	// Формируем файл для вывода.

	{

		$output = '<?xml version="1.0" encoding="UTF-8"?>';
		$output.= '<!DOCTYPE yml_catalog SYSTEM "shops.dtd">';
		$output.= '<yml_catalog date="2011-07-20 14:58">';
		$output.= '<shop>';
		$output.= '<name>'.SHOP_NAME.'</name>';
		$output.= '<company>'.SHOP_DESCRIPTION.'</company>';
		$output.= '<url>'.SHOP_URL.'</url>';
		$output.= '<currencies>';
		$output.= '<currency id="UAH" rate="1"/>';
		$output.= '</currencies>';
		$output.= '<categories>';

		foreach ($this->categories as $category_id => $category) {
			$output.= '<category id="'.$category_id.'" ';
			if  ($category['parent_id']!='') $output.= 'parentId="'.$category['parent_id'].'" ';
			$output.= '>'.$category['name'].'</category> ';
		}
		$output.= '</categories>';
		$output.= '<offers>';
		foreach ($outdata as $key => $data) {
			$output.= '<offer id="'.$key.'" available="'.(($data['stock_quantity']>0) ? 'true' : 'false').'">';
			$output.= '<url>'.$data['url'].'</url>';
			$output.= '<price>'.$data['price'].'</price>';
			$output.= '<currencyId>'.$data['currencyId'].'</currencyId>';
			$output.= '<categoryId>'.$data['categoryId'].'</categoryId>';
			foreach ($data['pictures'] as $picture) {
				$output.= '<picture>'.$picture.'</picture>';
			}
			$output.= '<vendor>'.$data['vendor'].'</vendor>';
			$output.= '<stock_quantity>'.$data['stock_quantity'].'</stock_quantity>';
			$output.= '<name>'.$data['name'].'</name>';
			$output.= '<description>'.$data['description'].'</description>';

			foreach ($data['params'] as $name => $value) {
				$output.= '<param name="'.$name.'">'.$value.'</param>';
			}

			$output.= '</offer>';
		
		}

		$output.= '</offers></shop></yml_catalog>';
		
		return (file_put_contents($file,$output)) ? count($outdata) : False;
		
	}

	private function getXLS($inputFileName)
	// Считываем файл $inputFileName. Возвращаем полученный массив данных
	{
		try {
			$inputFileType = PHPExcel_IOFactory::identify($inputFileName); // Определяем тип
			$this->WorkFileType=$inputFileType;
			$objReader = PHPExcel_IOFactory::createReader($inputFileType); // Создаем ридер
			$objReader->setReadDataOnly(true);
			$array = array();
			$worksheetNames = $objReader->listWorksheetNames($inputFileName); // Читаем имена страниц// Постранично читаем данные
			foreach ($worksheetNames as $wsName) {
				$arraysheet = array();
				$this->SheetNames[]=$wsName;
				$objReader->setLoadSheetsOnly($wsName);
				$oExcel = $objReader->load($inputFileName);
				$oExcel->setActiveSheetIndexByName($wsName);
				$aSheet = $oExcel->getActiveSheet();
				foreach ($aSheet->getRowIterator() as $rowId => $row) {
			   		$cellIterator = $row->getCellIterator();
		    		$item = array();
		    		$cellIterator->setIterateOnlyExistingCells(false);
		    		foreach($cellIterator as $cellId=>$cell) {
			       		array_push($item, $cell->getCalculatedValue());
			      	}
		    		array_push($arraysheet, $item);
			   	}
		    	array_push($array, $arraysheet);
		  	}
		  	unset($objReader);
		  	unset($inputFileType);
		  	return $array;
		} catch (Exception $exc) { die($exc->getMessage()); }
	}

}


$prg=new XMLtoYaml(INPUTFILE);
$result=$prg->getXMLData();

if (!$result) {
	die(sprintf("ВНИМАНИЕ !!! %s НЕ ОБРАБОТАН !!!\n",INPUTFILE));
} else {
	die(sprintf("%s ОБРАБОТАН. (Выгружено записей: %d)\n",INPUTFILE,$result));
}

?>