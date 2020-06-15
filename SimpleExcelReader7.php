<?php

/**
 * Class SimpleExcelReader7
 * only support microsoft excel 2007 or later version
 */

require_once "PHPExcel/PHPExcel.php";
class SimpleExcelReader7 {
    /*
     * Sheet Name
     */
    private $sheetNames = array();

    /*
     * Sheet Index, from zero
     */
    private $sheetIndexes = array(1);

    /*
     * the index of header row
     */
    private $headerRowIndex = 1;

    /*
     * the index of first data row
     */
    private $firstRowIndex = 2;
    private $rowReadCount = 0;
    private $filePath;


    function __construct($filePath, $options = []) {
        if(!$filePath || !file_exists($filePath)) {
            throw new Exception("the param \$filePath is not correct!");
        }
        $this->filePath = $filePath;
        $this->loadOptions($options);
    }

    public function load() {
        $zipClass = PHPExcel_Settings::getZipClass();
        $zip = new $zipClass;
        $zip->open($this->filePath);
        $sharedStrings = $this->getSharedStrings($zip);
        $sheetStyles = $this->getSheetStyles($zip);
        $worksheets = $this->getWorkSheetSelected($zip);
        $excel_data = array();
        foreach ($worksheets as $sheet) {
            $sd = simplexml_load_string($this->_getFromZipArchive($zip, 'xl/worksheets/'.$sheet['file']), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());
            $sheetData = $this->getSheetData($sd, $sharedStrings, $sheetStyles); //raw data
            $dimension = (string)$sd->dimension->attributes()['ref'];//A1:B5
            $range_from = explode(':', $dimension)[0];//A1;
            $from_aCoordinates = PHPExcel_Cell::coordinateFromString($range_from);//array('A','1');
            $range_to = explode(':', $dimension)[1];//B5
            $to_aCoordinates = PHPExcel_Cell::coordinateFromString($range_to);//array('B','5');
            $from_index_S = intval(PHPExcel_Cell::columnIndexFromString($from_aCoordinates[0]));//1
            $from_index_i = intval($from_aCoordinates[1]); //first row index
            $to_index_S = intval(PHPExcel_Cell::columnIndexFromString($to_aCoordinates[0]));//2
            $to_index_i = intval($to_aCoordinates[1]); //last row index

            $heads = array();//头部
            $data = array();
            $first_row_idx = 0;
            $col_count = 0;
            $col_letters = array();
            $row_to_select = $to_index_i;
            $max_col_index = $to_index_S;
            for($c = $from_index_S; $c <= $to_index_S; $c++){
                $columnLetter = PHPExcel_Cell::stringFromColumnIndex($c - 1);
                $col_letters[$c] = $columnLetter;
                $col_count++;
            }
            if(!$this->headerRowIndex || $this->headerRowIndex < $from_index_i || $this->headerRowIndex > $to_index_i){
                $this->headerRowIndex = 1;
            }
            if(!$this->firstRowIndex || $this->firstRowIndex > $to_index_i || $this->firstRowIndex <= $this->headerRowIndex){
                $this->firstRowIndex = $from_index_i + 1;
            }
            if($this->rowReadCount && $this->rowReadCount <= ($to_index_i - $this->firstRowIndex + 1)) {
                $row_to_select = $this->firstRowIndex + $this->rowReadCount - 1;
            }
            for($r = $from_index_i; $r <= $row_to_select; $r++){
                if(!($r == $this->headerRowIndex || $r >= $this->firstRowIndex)){
                    continue;
                }
                if(!$first_row_idx && $r > $from_index_i) $first_row_idx = $r;
                for($c = $from_index_S; $c <= $to_index_S; $c++){
                    if($max_col_index < $to_index_S && $c >= $max_col_index){
                        break;
                    }
                    if($r == $this->headerRowIndex){ //the head row
                        $h_v = $sheetData[$r][$col_letters[$c] . $r];
                        if($h_v && !in_array($h_v, $heads)){ //remove duplicates
                            $heads[] = $h_v;
                        }else{
                            $max_col_index = $c;
                            break;
                        }
                    }else{
                        $data[$r - $this->firstRowIndex][$c - $from_index_S] = $sheetData[$r][$col_letters[$c] . $r];
                    }
                }
            }
            $excel_data[$sheet['name']]['heads'] = $heads;
            $excel_data[$sheet['name']]['data'] = $data;
            $excel_data[$sheet['name']]['head_row_idx'] = $from_index_i;
            $excel_data[$sheet['name']]['first_row_idx'] = $first_row_idx;
            $excel_data[$sheet['name']]['col_count'] = $col_count;
            $excel_data[$sheet['name']]['col_letters'] = $col_letters;
            unset($sd);
        }
        $zip->close();
        return $excel_data;
    }
    private function getSheetData($sd, $sharedStrings, $sheetStyles){
        $sheetData = array();
        $rows = $sd->sheetData->row;
        foreach ($rows as $row) {
            $r_r = intval($row->attributes()['r']);
            $cols = $row->c;
            foreach ($cols as $col) {
                $c_r = (string)$col->attributes()['r'];
                $c_t = (string)$col->attributes()['t'];
                $c_s = (string)$col->attributes()['s'];
                $v = (string)$col->v;
                if(isset($v) && $v != null) {
                    $v = $this->getRealValue($v, $sharedStrings, $sheetStyles, $c_t, $c_s);
                }
                $sheetData[$r_r][$c_r] = $v;
            }
        }
        return $sheetData;
    }
    private function getRealValue($v, $sharedStrings, $sheetStyles, $c_t, $c_s){
        if(!$c_t && !$c_s){
            return $v;
        }
        if ($c_t == 's') {
            return trim($sharedStrings[intval($v)]);
        } elseif ($c_s) {
            $format_code = $sheetStyles[intval($c_s)];
            if($format_code){
                $format_code = "~" . $format_code;
                if(strpos($format_code, "yy") && strpos($format_code, "m") && strpos($format_code, "d")){
                    if(strpos($format_code, "h:m")){
                        if(strpos($format_code, ":s")){
                            return gmdate('Y-m-d H:i:s', PHPExcel_Shared_Date::ExcelToPHP($v));
                        }
                        return gmdate('Y-m-d H:i', PHPExcel_Shared_Date::ExcelToPHP($v));
                    }
                    return gmdate('Y-m-d', PHPExcel_Shared_Date::ExcelToPHP($v));
                }elseif(strpos($format_code, "h:m")){
                    if(strpos($format_code, ":s")){
                        return gmdate('H:i:s', PHPExcel_Shared_Date::ExcelToPHP($v));
                    }
                    return gmdate('H:i', PHPExcel_Shared_Date::ExcelToPHP($v));
                }elseif(strpos($format_code, "@")){
                    return $v;
                }
            }
        }
        if(is_numeric($v)){
            if(strpos($v, ".") && !strpos($v, "E")){
                return round($v, 4);
            }elseif(strpos($v, "E-")){
                return round(number_format($v, 8, '.', ''), 4);
            }
            return number_format($v, 0, '', '');
            /*$v = number_format($v, 4, '.', '');
            if(preg_match("/^\d+\.[0]+$/", $v, $match)){
                return number_format($v, 0, '', '');
            }*/
        }
        return $v;
    }
    private function getSharedStrings($zip){
        $sharedStrings = array();
        $ws = simplexml_load_string($this->_getFromZipArchive($zip, 'xl/sharedStrings.xml'), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());
        foreach ($ws->si as $s) {
            $str = $s->t;
            $str_r = $s->r;
            $v = "";
            if($str){
                $v = (string)$str;
            }elseif($str_r){
                foreach($str_r as $r){
                    $v .= (string)$r->t;
                }
            }
            $sharedStrings[] = $v;
        }
        return $sharedStrings;
    }
    private function getSheetStyles($zip){
        $sheetStyles = array();
        $ss = simplexml_load_string($this->_getFromZipArchive($zip, 'xl/styles.xml'), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());
        $fmts = $ss->numFmts->numFmt;
        $format_code_list = array();
        if($fmts) {
            foreach ($fmts as $f) {
                $format_code_list[(string)$f->attributes()["numFmtId"]] = (string)$f->attributes()["formatCode"];
            }
        }
        $formats_others = $this->getOtherFormat();
        $cf = $ss->cellXfs->xf;
        if($cf) {
            foreach ($cf as $c) {
                $nid = (string)$c->attributes()["numFmtId"];
                $fc = null;
                if ($nid && intval($nid) < 164) {
                    $fc = $formats_others[intval($nid)];
                } elseif ($nid) {
                    $fc = $format_code_list[$nid];
                }
                $sheetStyles[] = $fc;
            }
        }
        return $sheetStyles;
    }
    private function getOtherFormat(){
        $ret = array();
        $ret[0] = PHPExcel_Style_NumberFormat::FORMAT_GENERAL;
        $ret[1] = '0';
        $ret[2] = '0.00';
        $ret[3] = '#,##0';
        $ret[4] = '#,##0.00';

        $ret[9] = '0%';
        $ret[10] = '0.00%';
        $ret[11] = '0.00E+00';
        $ret[12] = '# ?/?';
        $ret[13] = '# ??/??';
        $ret[14] = 'mm-dd-yy';
        $ret[15] = 'd-mmm-yy';
        $ret[16] = 'd-mmm';
        $ret[17] = 'mmm-yy';
        $ret[18] = 'h:mm AM/PM';
        $ret[19] = 'h:mm:ss AM/PM';
        $ret[20] = 'h:mm';
        $ret[21] = 'h:mm:ss';
        $ret[22] = 'm/d/yy h:mm';

        $ret[37] = '#,##0 ;(#,##0)';
        $ret[38] = '#,##0 ;[Red](#,##0)';
        $ret[39] = '#,##0.00;(#,##0.00)';
        $ret[40] = '#,##0.00;[Red](#,##0.00)';

        $ret[44] = '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)';
        $ret[45] = 'mm:ss';
        $ret[46] = '[h]:mm:ss';
        $ret[47] = 'mmss.0';
        $ret[48] = '##0.0E+0';
        $ret[49] = '@';

        // CHT
        $ret[27] = '[$-404]e/m/d';
        $ret[30] = 'm/d/yy';
        $ret[36] = '[$-404]e/m/d';
        $ret[50] = '[$-404]e/m/d';
        $ret[57] = '[$-404]e/m/d';

        // THA
        $ret[59] = 't0';
        $ret[60] = 't0.00';
        $ret[61] = 't#,##0';
        $ret[62] = 't#,##0.00';
        $ret[67] = 't0%';
        $ret[68] = 't0.00%';
        $ret[69] = 't# ?/?';
        $ret[70] = 't# ??/??';
        return $ret;
    }
    private function getWorkSheetSelected($zip) {
        $worksheets = $this->getWorksheets($zip);
        $retSheets = array();
        if(count($this->sheetNames)) {
            foreach ($worksheets as $idx => $ws) {
                if(in_array($ws['name'], $this->sheetNames)){
                    $retSheets[] = $ws;
                }
            }
        }elseif(count($this->sheetIndexes)) {
            foreach ($worksheets as $idx => $ws) {
                if(in_array($idx, $this->sheetIndexes)){
                    $retSheets[] = $ws;
                }
            }
        }
        return $retSheets;
    }

    private function getWorksheets($zip){
        $sheets = array();
        $rs = simplexml_load_string($this->_getFromZipArchive($zip, 'xl/_rels/workbook.xml.rels'), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());
        $ralation_ships = $rs->Relationship;
        $ralations = array();
        foreach ($ralation_ships as $r) {
            $target = (string)$r->attributes()['Target'];
            $type = (string)$r->attributes()['Type'];
            $id = (string)$r->attributes()['Id'];
            if($type == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'){
                $ralations[$id] = str_replace('worksheets/', '', $target);
            }
        }
        $xml_str = str_replace('r:id', 'rId', $this->_getFromZipArchive($zip, 'xl/workbook.xml'));
        $ws = simplexml_load_string($xml_str, 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());
        foreach ($ws->sheets->sheet as $s) {
            $rid = (string)$s['rId'];
            $name = (string)$s['name'];
            $sheet_file = $ralations[$rid];
            $sheets[(string)$s['sheetId']] = array('Id'=>(string)$s['sheetId'], 'name'=>$name, 'file' => $sheet_file, 'rid' => $rid);
        }
        return $sheets;
    }

    private function _getFromZipArchive($archive, $fileName = '')
    {
        if (strpos($fileName, '//') !== false){
            $fileName = substr($fileName, strpos($fileName, '//') + 1);
        }
        $fileName = PHPExcel_Shared_File::realpath($fileName);
        // Apache POI fixes
        $contents = $archive->getFromName($fileName);
        if ($contents === false){
            $contents = $archive->getFromName(substr($fileName, 1));
        }
        return $contents;
    }

    private function loadOptions($opts){
        if(isset($opts["sheet_names"])) {
            if(is_array($opts["sheet_names"])) {
                $this->sheetNames = $opts["sheet_names"];
            }else{
                $this->sheetNames = [$opts["sheet_names"]];
            }
        }
        if(isset($opts["sheet_indexes"])) {
            if(is_array($opts["sheet_indexes"])) {
                $this->sheetIndexes = $opts["sheet_indexes"];
            }else{
                $this->sheetIndexes = [$opts["sheet_indexes"]];
            }
        }
        if(isset($opts["header_row_index"])) {
            $this->headerRowIndex = $opts["header_row_index"];
        }
        if(isset($opts["first_row_index"])) {
            $this->firstRowIndex = $opts["first_row_index"];
        }
        if(isset($opts["row_read_count"])) {
            $this->rowReadCount = $opts["row_read_count"];
        }
    }
}