<?php

/*
	PHP Tutorial: Office XML - Create an Excel Document
	Copyright 2010 CS Truter
	Created by: Christoff Truter	
	Url: http://www.cstruter.com
	Email: christoff@cstruter.com
*/

class Excel
{
	public $node;
	protected $ss = "urn:schemas-microsoft-com:office:spreadsheet";
}

class Workbook extends Excel
{	
	public function __construct()
	{
		$xml = '<?mso-application progid="Excel.Sheet"?>'.
				'<ss:Workbook xmlns:ss="'.$this->ss.'" />';
		$this->node = new SimpleXMLElement($xml);
	}
	
	public function addWorksheet($name)
	{
		return new Worksheet($name, $this->node);
	}
	
	public function Output($filename = NULL)
	{
		if (!isset($filename))
		{
			header("content-type: text/xml");
			echo $this->node->asXml();
		}
		else
		{
			file_put_contents($filename, $this->node->asXml());
		}
	}
}

class Worksheet extends Excel
{
	public function __construct($name, $node)
	{
		$this->node = $node->addChild('ss:Worksheet');
		$this->node->addAttribute('ss:Name', $name, $this->ss);
	}
	
	public function addTable(array $values = NULL)
	{
		return new Table($this->node, $values);
	}
}

class Table extends Excel
{
	public function __construct($node, $values)
	{
		$this->node = $node->addChild('ss:Table');
		if ($values != NULL)
		{	
			$this->addRows($values);
		}
	}
	
	public function addRows(array $values = NULL)
	{				
		foreach($values as $key=>$cells)
		{
			$this->addRow($cells);
		}	
	}
	
	public function addRow(array $values = NULL)
	{
		return new Row($this->node, $values);
	}
}

class Row extends Excel
{
	public function __construct($node, $values)
	{
		$this->node = $node->addChild('ss:Row');
		if ($values != NULL)
		{
			$this->addCells($values);
		}
	}
	
	public function addCells(array $values = NULL)
	{
		foreach($values as $value)
		{
			$this->addCell($value);
		}
	}
	
	public function addCell($value)
	{
		return new Cell($value, $this->node);
	}
}

class Cell  extends Excel
{
	public function __construct($value, $node)
	{
		$this->node = $node->addChild('ss:Cell');
		$Data = $this->node->addChild('ss:Data', $value, $this->ss);
		$Data->addAttribute('ss:Type', 'String', $this->ss);
	}
}

?>