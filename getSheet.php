<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Excel2007;

$spreadsheet = new Spreadsheet();
$spreadsheet->getDefaultStyle()->getFont()->setSize(13);
$sheet = $spreadsheet->getActiveSheet();

$message = $_POST["msg"];
$default_hours = $_POST["hrs"];
$lines = explode("\n", $message); //divide message into lines

class Entry{
    public function __construct($f_n, $l_n, $h){
        $this->first_name = $f_n;
        $this->last_name = $l_n;
        $this->hours = $h;
    }
    public $first_name = "MISSING";
    public $last_name = "MISSING";
    public $hours = "";
}

function compare_entries($a, $b){ //compare entries by last name
    return strcmp($a->last_name, $b->last_name);
}

$entries = array();

for($i = 0; $i < sizeof($lines); $i++){ //for each line
    $hours = $default_hours;
    $names = explode(" ", $lines[$i]); //divide name by space

    if(sizeof($names) < 2) continue;

    $names[sizeof($names)-1] = trim($names[sizeof($names)-1]);

    if(is_numeric($names[sizeof($names)-1])){ //if last element is numeric
        $hours = $names[sizeof($names)-1]; //set hours to last element
        array_pop($names); //remove hours from name array
    }

    if(sizeof($names) < 2) continue;

    array_push($entries, new Entry($names[0], $names[sizeof($names)-1], $hours)); //add new entry to array
}

usort($entries, "compare_entries"); //sort array by last name

$sheet->setCellValue('A1', 'First name');
$sheet->setCellValue('B1', 'Last name');
$sheet->setCellValue('C1', 'Hours');

$sheet->getStyle('A1')->getFont()->setBold(true);
$sheet->getStyle('B1')->getFont()->setBold(true);
$sheet->getStyle('C1')->getFont()->setBold(true);

for($i = 0; $i < sizeof($entries); $i++){
    $current_entry = $entries[$i];
    $sheet->setCellValue('A'.($i+2), $current_entry->first_name);
    $sheet->setCellValue('B'.($i+2), $current_entry->last_name);
    $sheet->setCellValue('C'.($i+2), $current_entry->hours);
}

$sheet->getColumnDimension('A')->setAutoSize(true);
$sheet->getColumnDimension('B')->setAutoSize(true);
$sheet->getColumnDimension('C')->setAutoSize(true);

$writer = new Excel2007($spreadsheet);
// Redirect output to a client’s web browser (Xlsx)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="'.$_POST["name"].' '.$_POST["date"].' hours.xlsx"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');
// If you're serving to IE over SSL, then the following may be needed
header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header('Pragma: public'); // HTTP/1.0

$writer->save('php://output');
