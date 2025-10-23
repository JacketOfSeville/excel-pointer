# excel-pointer

ExcelPointer is a PHP library for navigating Excel sheets, designed to work with PHPSpreadsheet or similar tools. It supports both XLS and XLSX boundaries and provides movement and coordinate tracking.

## Installation

Install via Composer:

```bash
composer require jacketofseville/excel-pointer
```


## Basic Usage

```php
require 'vendor/autoload.php';

$pointer = new ExcelPointer(); // Defaults to XLSX boundaries
$pointer->right();      // Move one column right
$pointer->down(5);      // Move five rows down
$pointer->left(2);      // Move two columns left
$pointer->up();         // Move one row up

// Get current coordinate as string (e.g., 'C5')
$coord = $pointer->coord();

// Get current coordinate as array
$coordArr = $pointer->coord('array'); // ['column' => 3, 'row' => 5]

// Get boundary (max column/row reached so far)
$boundary = $pointer->boundary();
```

## More Examples

### Using tab() for horizontal navigation

The `tab()` method returns the current coordinate and then moves the pointer one column to the right. This is useful for filling a row, like tabbing through cells in Excel:

```php
$pointer = new ExcelPointer();
for ($i = 0; $i < 5; $i++) {
	$cell = $pointer->tab(); // Returns current cell, then moves right
	echo "Writing to $cell\n";
	// $sheet->setCellValue($cell, "Value $i");
}
```

### Filling a table row with tab()

```php
$pointer = new ExcelPointer();
$pointer->down(3); // Move to row 4
for ($i = 0; $i < 10; $i++) {
	$cell = $pointer->tab();
	// $sheet->setCellValue($cell, "Row 4, Col $i");
}
```

### Using tab() with PHPSpreadsheet

```php
$pointer = new ExcelPointer();
foreach (["Name", "Email", "Phone"] as $value) {
	$cell = $pointer->tab();
	$sheet->setCellValue($cell, $value);
}
$pointer->enter(); // Move to next row, first column
foreach (["Alice", "alice@example.com", "555-1234"] as $value) {
	$cell = $pointer->tab();
	$sheet->setCellValue($cell, $value);
}
```

### Resetting to first column of next row

```php
$pointer->enter(); // Moves to first column of next row
```

### Checking boundaries before moving

```php
try {
	$pointer->right(20000); // Throws if out of bounds
} catch (OutOfBoundsException $e) {
	echo $e->getMessage();
}
```

## XLS vs XLSX Boundaries

By default, ExcelPointer uses XLSX boundaries (16384 columns, 1048576 rows). To use legacy XLS limits (256 columns, 65536 rows):

```php
$pointer = new ExcelPointer('xls');
```

## Movement Methods

- `right($n = 1)`: Move right by $n columns
- `left($n = 1)`: Move left by $n columns
- `down($n = 1)`: Move down by $n rows
- `up($n = 1)`: Move up by $n rows
- `enter()`: Move to the first column of the next row
- All movement methods check bounds and throw exceptions if the move is invalid.

## Error Handling

If you attempt to move out of bounds, an `OutOfBoundsException` is thrown:

```php
try {
	$pointer->right(20000); // Too far for XLSX
} catch (OutOfBoundsException $e) {
	echo $e->getMessage();
}
```

If you pass an invalid parameter (e.g., negative or zero), an `InvalidArgumentException` is thrown.

## Testing

Basic tests are provided in `tests/ExcelPointerTest.php`. Run them with:

```bash
php tests/ExcelPointerTest.php
```

## Integration with PHPSpreadsheet

ExcelPointer is designed to help you navigate cell coordinates and boundaries. You can use its output to set or get values in PHPSpreadsheet:

```php
$sheet->setCellValue($pointer->coord(), 'Value');
$pointer->right();
$sheet->setCellValue($pointer->coord(), 'Next Value');
```

## License

MIT
