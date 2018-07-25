# EasySpreadsheets

<a href="https://packagist.org/packages/natanaugusto/easyspreadsheets"><img src="https://poser.pugx.org/natanaugusto/easyspreadsheets/d/total.svg" alt="Total Downloads"></a>
<a href="https://packagist.org/packages/natanaugusto/easyspreadsheets"><img src="https://poser.pugx.org/natanaugusto/easyspreadsheets/v/stable.svg" alt="Latest Stable Version"></a>
<a href="https://packagist.org/packages/natanaugusto/easyspreadsheets"><img src="https://poser.pugx.org/natanaugusto/easyspreadsheets/license.svg" alt="License"></a>
<a href="https://www.flyingdonut.io/app/project/project-id=5b57c6e7e4b015ad58e36c12"><img src="https://www.flyingdonut.io/api/projects/5b57c6e7e4b015ad58e36c12/iterations/current/status.svg" alt="Flying Donut"></a>

That's a easy way to use [PhpSpreadsheet](https://phpspreadsheet.readthedocs.io).

What you can do with this:
 - Load a spreadsheet very easily.
 - Read spreadsheets with header
 - Read cells colors
 - Read rows associated with their headers
 - Paint cells and texts
 - That's it for now

Install:
```shell
composer require natanaugusto/easyspreadsheets
```
Begins:
```php
require_once('../vendor/autoload.php');

use EasySpreadsheets\Handler as EasySpreadsheet;
$file = dirname(__FILE__) . DIRECTORY_SEPARATOR . 'example.xlsx';

$spread = new EasySpreadsheet;
$spread->load($file);
```

Get all rows:
```php
$spread->getRows()
```

Get the current row:
```php
$spread->getRow()
```

Get a especific row:
```php
$spread->getRow(2)
```
(Yes, my English is a very bad shit)
