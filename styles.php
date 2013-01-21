<?php

/*
	PHP Tutorial: Office XML - Create an Excel Document - Styles
	Copyright 2010 CS Truter
	Created by: Christoff Truter	
	Url: http://www.cstruter.com
	Email: christoff@cstruter.com
*/

class Styles extends Excel
{
	public function __construct($Workbook)
	{
		$this->node = $Workbook->node->addChild('ss:Styles');
	}
	
	public function setStyle($Id, $element)
	{
		$element->node->addAttribute('ss:StyleID', $Id, $this->ss);
		return new Style($this->node, $Id);
	}	
}

class Style extends Excel
{
	private $font;
	
	public function __construct($node, $Id)
	{
		$this->node = $node->addChild('ss:Style');
		$this->node->addAttribute('ss:ID', $Id, $this->ss);
	}
	
	public function setFont($fontname, $size, $colour, $bold)
	{
		if (!isset($this->font)) 
		{
			$this->font = $this->node->addChild('ss:Font');
			
			if (isset($fontname)) {
				$this->font->addAttribute('ss:FontName', $fontname, $this->ss);
			}
			if (isset($size)) {
				$this->font->addAttribute('ss:Size', $size, $this->ss);
			}
			if (isset($colour)) {
				$this->font->addAttribute('ss:Color', $colour, $this->ss);
			}
			if ($bold) {
				$this->font->addAttribute('ss:Bold', '1', $this->ss);
			}
		}
	}
}

?>