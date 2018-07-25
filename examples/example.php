<?php
require_once('vendor/autoload.php');

use EasySpreadsheets\Handler as EasySpreadsheet;
$file = dirname(__FILE__) . DIRECTORY_SEPARATOR . 'example.xlsx';

$spread = new EasySpreadsheet;
$spread->load($file);

echo "Get all rows\n";
print_r($spread->getRows());

echo "Get the current row\n";
print_r($spread->getRow());

echo "Get a specific row\n";
print_r($spread->getRow(4));