# SimpleExcelReader7

An excel reader based on PHPExcel

# example

```
require_once "SimpleExcelReader7.php";
$excelReader = new SimpleExcelReader7("test.xlsx");
$excelData = $excelReader->load();
var_dump($excelData);
```
