<?php
require_once __DIR__ . '/../src/ExcelPointer.php';
require_once __DIR__ . '/../src/ExcelColumn.php';

use JacketOfSeville\ExcelPointer\ExcelPointer;

function assertEqual($a, $b, $msg = '') {
    if ($a !== $b) {
        echo "Assertion failed: $msg\n";
        echo "Expected: "; var_export($b); echo "\n";
        echo "Got: "; var_export($a); echo "\n";
        exit(1);
    }
}

function test_basic_movement() {
    $p = new ExcelPointer();
    assertEqual($p->getRow(), 1, 'Initial row');
    assertEqual($p->getColumn(), 1, 'Initial column');
    $p->right(5);
    assertEqual($p->getColumn(), 6, 'Moved right 5');
    $p->down(3);
    assertEqual($p->getRow(), 4, 'Moved down 3');
    $p->left(2);
    assertEqual($p->getColumn(), 4, 'Moved left 2');
    $p->up(1);
    assertEqual($p->getRow(), 3, 'Moved up 1');
}

function test_bounds_xlsx() {
    $p = new ExcelPointer('xlsx');
    try {
        $p->right(16384); 
        assertEqual(true, false, 'Should throw on right out of bounds');
    } catch (OutOfBoundsException $e) {}
    try {
        $p->down(1048576); 
        assertEqual(true, false, 'Should throw on down out of bounds');
    } catch (OutOfBoundsException $e) {}
}

function test_bounds_xls() {
    $p = new ExcelPointer('xls');
    try {
        $p->right(256); 
        assertEqual(true, false, 'Should throw on right out of bounds (xls)');
    } catch (OutOfBoundsException $e) {}
    try {
        $p->down(65536); 
        assertEqual(true, false, 'Should throw on down out of bounds (xls)');
    } catch (OutOfBoundsException $e) {}
}

function test_invalid_params() {
    $p = new ExcelPointer();
    try {
        $p->right(0);
        assertEqual(true, false, 'Should throw on right(0)');
    } catch (InvalidArgumentException $e) {}
    try {
        $p->down(-1);
        assertEqual(true, false, 'Should throw on down(-1)');
    } catch (InvalidArgumentException $e) {}
}

test_basic_movement();
test_bounds_xlsx();
test_bounds_xls();
test_invalid_params();
echo "All tests passed.\n";
