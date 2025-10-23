<?php
namespace JacketOfSeville\ExcelPointer;

use InvalidArgumentException;
use OutOfBoundsException;

class ExcelPointer
{
    private const MAX_COLUMN = 16384;
    private const MAX_ROW = 1048576;
    private const MIN_COLUMN = 1;
    private const MIN_ROW = 1;

    /** @var int */
    private $column;

    /** @var int */
    private $row;

    /** @var string */
    private $coordinate;

    /** @var int */
    private $maxColumn;

    /** @var int */
    private $maxRow;

    /** @var int */
    private $maxColumnLimit;

    /** @var int */
    private $maxRowLimit;

    /**
     * ExcelPointer constructor.
     * Initializes the pointer to the top-left cell and sets initial boundaries.
     *
     * @param string $format Either 'xlsx' (default) or 'xls' to set appropriate limits
     */
    public function __construct($format = 'xlsx')
    {
        $this->column = self::MIN_COLUMN;
        $this->row = self::MIN_ROW;
        $this->maxColumn = self::MIN_COLUMN;
        $this->maxRow = self::MIN_ROW;
        // Set format-specific limits
        switch (strtolower($format)) {
            case 'xls':
                // XLS (BIFF8) limits: 256 columns (IV), 65536 rows
                $this->maxColumnLimit = 256;
                $this->maxRowLimit = 65536;
                break;
            case 'xlsx':
            default:
                // XLSX limits: 16384 columns (XFD), 1048576 rows
                $this->maxColumnLimit = self::MAX_COLUMN;
                $this->maxRowLimit = self::MAX_ROW;
                break;
        }

        $this->updateCoordinate();
    }

    /**
     * Update the cached coordinate string and adjust observed max column/row.
     *
     * @return void
     */
    private function updateCoordinate()
    {
        $this->coordinate = ExcelColumn::coord($this->column, $this->row);
        $this->maxColumn = max($this->maxColumn, $this->column);
        $this->maxRow = max($this->maxRow, $this->row);
    }

    /**
     * Move the pointer right by a given number of columns.
     *
     * @param int $n Number of columns to move (default 1)
     * @return self
     * @throws OutOfBoundsException if the movement would exceed the max column
     */
    public function right($n = 1)
    {
        if (!is_int($n) || $n < 1) {
            throw new InvalidArgumentException('Parameter $n must be a positive integer');
        }
        if ($this->column + $n > $this->maxColumnLimit) {
            throw new OutOfBoundsException(sprintf('Column right limit reached (%s | %d)', ExcelColumn::column($this->maxColumnLimit), $this->maxColumnLimit));
        }
        $this->column += $n;
        $this->updateCoordinate();
        return $this;
    }

    /**
     * Move the pointer left by a given number of columns.
     *
     * @param int $n Number of columns to move (default 1)
     * @return self
     * @throws OutOfBoundsException if the movement would go before the first column
     */
    public function left($n = 1)
    {
        if (!is_int($n) || $n < 1) {
            throw new InvalidArgumentException('Parameter $n must be a positive integer');
        }
        if ($this->column - $n < self::MIN_COLUMN) {
            throw new OutOfBoundsException(sprintf('Column left limit reached (%s | %d)', ExcelColumn::column(self::MIN_COLUMN), self::MIN_COLUMN));
        }
        $this->column -= $n;
        $this->updateCoordinate();
        return $this;
    }

    /**
     * Move the pointer down by a given number of rows.
     *
     * @param int $n Number of rows to move (default 1)
     * @return self
     * @throws OutOfBoundsException if the movement would exceed the max row
     */
    public function down($n = 1)
    {
        if (!is_int($n) || $n < 1) {
            throw new InvalidArgumentException('Parameter $n must be a positive integer');
        }
        if ($this->row + $n > $this->maxRowLimit) {
            throw new OutOfBoundsException(sprintf('Row lower limit reached (%d)', $this->maxRowLimit));
        }
        $this->row += $n;
        $this->updateCoordinate();
        return $this;
    }

    /**
     * Move the pointer up by a given number of rows.
     *
     * @param int $n Number of rows to move (default 1)
     * @return self
     * @throws OutOfBoundsException if the movement would go before the first row
     */
    public function up($n = 1)
    {
        if (!is_int($n) || $n < 1) {
            throw new InvalidArgumentException('Parameter $n must be a positive integer');
        }
        if ($this->row - $n < self::MIN_ROW) {
            throw new OutOfBoundsException(sprintf('Row upper limit reached (%d)', self::MIN_ROW));
        }
        $this->row -= $n;
        $this->updateCoordinate();
        return $this;
    }

    /**
     * Move the pointer to the first column of the next row (like pressing Enter).
     *
     * @return self
     * @throws OutOfBoundsException if the pointer would exceed the max row
     */
    public function enter()
    {
        if ($this->row >= $this->maxRowLimit) {
            throw new OutOfBoundsException(sprintf('Row lower limit reached (%d)', $this->maxRowLimit));
        }
        $this->row++;
        $this->column = self::MIN_COLUMN;
        $this->updateCoordinate();
        return $this;
    }

    /**
     * Get the current coordinate.
     *
     * @param string $mode 'string' or 'array'
     * @return string|array
     * @throws InvalidArgumentException
     */
    public function coord(string $mode = 'string')
    {
        switch ($mode) {
            case 'string':
                return $this->coordinate;
            case 'array':
                return ['column' => $this->column, 'row' => $this->row];
            default:
                throw new InvalidArgumentException("Invalid mode: {$mode}");
        }
    }

    /**
     * Returns the current coordinate and moves the pointer to the next cell to the right.
     *
     * @return string The coordinate of the cell before the move.
     */
    public function tab()
    {
        $ret = $this->coordinate;
        $this->right();
        return $ret;
    }

    /**
     * Get the boundary coordinate (max column/row seen).
     *
     * @param string $mode 'string' or 'array'
     * @return string|array
     * @throws InvalidArgumentException
     */
    public function boundary(string $mode = 'array')
    {
        switch ($mode) {
            case 'string':
                return ExcelColumn::coord($this->maxColumn, $this->maxRow);
            case 'array':
                return ['column' => $this->maxColumn, 'row' => $this->maxRow];
            default:
                throw new InvalidArgumentException("Invalid mode: {$mode}");
        }
    }

    /**
     * Get the current row index (1-based).
     *
     * @return int
     */
    public function getRow()
    {
        return $this->row;
    }

    /**
     * Get the current column index.
     * 
     * @param string $mode 'int' for integer index, 'str' for Excel column string
     * @return int|string
     */
    public function getColumn($mode = 'int')
    {
        switch ($mode) {
            case 'int':
                return $this->column;
            case 'str':
                return ExcelColumn::column($this->column);
            default:
                throw new InvalidArgumentException("Invalid mode: {$mode}");
        }
    }
}
