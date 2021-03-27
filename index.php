<meta charset="Windows-1251">

<?php

class Excel_XML
{

    /**
     * Header of excel document (prepended to the rows)
     * 
     * Copied from the excel xml-specs.
     * 
     * @access private
     * @var string
     */
    var $header = "<?xml version=\"1.0\" encoding=\"UTF-8\"?\>
<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"
 xmlns:x=\"urn:schemas-microsoft-com:office:excel\"
 xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"
 xmlns:html=\"http://www.w3.org/TR/REC-html40\">";

    /**
     * Footer of excel document (appended to the rows)
     * 
     * Copied from the excel xml-specs.
     * 
     * @access private
     * @var string
     */
    var $footer = "</Workbook>";

    /**
     * Document lines (rows in an array)
     * 
     * @access private
     * @var array
     */
    var $lines = array ();

    /**
     * Worksheet title
     *
     * Contains the title of a single worksheet
     *
     * @access private 
     * @var string
     */
    var $worksheet_title = "Table1";

    /**
     * Add a single row to the $document string
     * 
     * @access private
     * @param array 1-dimensional array
     * @todo Row-creation should be done by $this->addArray
     */
    function addRow ($array)
    {

        // initialize all cells for this row
        $cells = "";

        // foreach key -> write value into cells
        foreach ($array as $k => $v):

            if(function_exists("iconv")) {
				$cells .= "<Cell><Data ss:Type=\"String\">" . iconv("gbk", "utf-8", $v) . "</Data></Cell>\n"; // by steven.2008-1-17.
			}elseif(function_exists("mb_convert_encoding")){
				$cells .= "<Cell><Data ss:Type=\"String\">" . mb_convert_encoding($v, "utf-8", "gbk") . "</Data></Cell>\n"; // by steven.2008-1-17.
			}else{
				$cells .= "<Cell><Data ss:Type=\"String\">" . utf8_encode($v) . "</Data></Cell>\n"; 
			}
		endforeach;
        // transform $cells content into one row
        $this->lines[] = "<Row>\n" . $cells . "</Row>\n";

    }

    /**
     * Add an array to the document
     * 
     * This should be the only method needed to generate an excel
     * document.
     * 
     * @access public
     * @param array 2-dimensional array
     * @todo Can be transfered to __construct() later on
     */
    function addArray ($array)
    {

        // run through the array and add them into rows
        foreach ($array as $k => $v):
            $this->addRow ($v);
        endforeach;

    }

    /**
     * Set the worksheet title
     * 
     * Checks the string for not allowed characters (:\/?*),
     * cuts it to maximum 31 characters and set the title. Damn
     * why are not-allowed chars nowhere to be found? Windows
     * help's no help...
     *
     * @access public
     * @param string $title Designed title
     */
    function setWorksheetTitle ($title)
    {

        // strip out special chars first
        $title = preg_replace ("/[\\\|:|\/|\?|\*|\[|\]]/", "", $title);

        // now cut it to the allowed length
        $title = substr ($title, 0, 31);

        // set title
        $this->worksheet_title = $title;

    }

    /**
     * Generate the excel file
     * 
     * Finally generates the excel file and uses the header() function
     * to deliver it to the browser.
     * 
     * @access public
     * @param string $filename Name of excel file to generate (...xls)
     */
    function generateXML ($filename)
    {

        // deliver header (as recommended in php manual)
        header("Content-Type: application/vnd.ms-excel; charset=UTF-8");
        header("Content-Disposition: inline; filename=\"" . $filename . ".xls\"");

        // print out document to the browser
        // need to use stripslashes for the damn ">"
        echo stripslashes ($this->header);
        echo "\n<Worksheet ss:Name=\"" . $this->worksheet_title . "\">\n<Table>\n";
        echo "<Column ss:Index=\"1\" ss:AutoFitWidth=\"0\" ss:Width=\"110\"/>\n";
        echo implode ("\n", $this->lines);
        echo "</Table>\n</Worksheet>\n";
        echo $this->footer;

    }

}



error_reporting(-1);

header('Content-Type: text/html; charset=Windows-1251');

mb_internal_encoding("Windows-1251");
function Parse($String, $p1, $p2) {
	//echo $String;
	$num1 = strpos($String, $p1);
	if ($num1 === false) return 0;
	$x=0;
	$num2 = substr($String, $num1);
	$str[$x] = substr($num2, 0, strlen($p2)+strpos($num2, $p2));
	$str[$x]=str_replace($p1,'',$str[$x]);
	$str[$x]=str_replace($p2,'',$str[$x]);
	//echo $str[$x];
	while($num1 != false){ 
	//$num1 = strlen($p2)+strpos($num2, $p2);
	$num2 = substr($num2, strlen($p2));
	$num1 = strpos($num2, $p1);
	if($num1===false) break;
	$num2 = substr($num2, $num1);
	$x++;
	$str[$x] = substr($num2, 0, strlen($p2)+strpos($num2, $p2));
	$str[$x]=str_replace($p1,'',$str[$x]);
	$str[$x]=str_replace($p2,'',$str[$x]);
	//echo $str[$x];
	}
	return $str;
}

$String = file_get_contents('https://mebelmassive77.ru/yml/yandex.php?pas=fylhtq');
$categories = Parse($String, '"><![CDATA[', ']]></category>');
$content = simplexml_load_file('https://mebelmassive77.ru/yml/yandex.php?pas=fylhtq');
$content = $content->shop->categories->category;
$x=0;
foreach($content as $content){
	$cont[$x]=$content["id"];
	$x++;
}

$content = array_fill_keys($cont, '0');
$content = array_keys($content);
//var_dump($content);
$x=0;
foreach($content as $cont){
 $content[$cont]=$categories[$x];
 $x++;
}
//var_dump($content[1]);

//$categories = Parse($String, '"><![CDATA[', ']]></category>');
$categoryId = Parse($String, '<categoryId>', '</categoryId>');
$name = Parse($String, '<name><![CDATA[', ']]></name>');
$description = Parse($String, '<description><![CDATA[', ']]></description>');
$price = Parse($String, '<price>', '</price>');
$picture = Parse($String, '<picture>', '</picture>');
$pickup = Parse($String, '<pickup>', '</pickup>');
$store = Parse($String, '<store>', '</store>');
$x=0;

$file="demo.xls";

header("Content-Disposition: attachment; filename=$file");
header("Connection: Keep-Alive");
header("Content-Type: application/vnd.ms-excel");


echo '<table border="1">';
	echo '<tr>'.'<td>Категория</td>'.'<td>Название</td>'.'<td>Описание</td>'.'<td>Цена</td>'.'<td>Фото</td>'.'<td>Популярный товар</td>'.'<td>В наличии</td>'.'</tr>';
foreach($categoryId as $categoryId){
	if($pickup[$x]=="false") $pickup[$x]="Нет";
	if($pickup[$x]=="true") $pickup[$x]="Да";
	if($store[$x]=="false") $store[$x]="Нет";
	if($store[$x]=="true") $store[$x]="Да";
	echo '<tr>';
	echo '<td>'.$content[$categoryId].'</td>'.'<td>'.$name[$x].'</td>'.'<td>'.strip_tags($description[$x]).'</td>'.'<td>'.$price[$x].'</td>'.'<td>'.$picture[$x].'</td>'.'<td>'.$pickup[$x].'</td>'.'<td>'.$store[$x].'</td>';
	$x++;
	echo '</tr>';
}
echo '</table>';


//$String = Parse($String, 'id', 'available');
//echo $categoryId;