<?php
class AlxdExportXLSX
{
    const ZERO_TIMESTAMP = 2209161600;
    const SEC_IN_DAY = 86400;

    public $templateFile = 'template.xlsx';
    public $exportDir = '/var/www/exports/xlsx';

    private $exportFile = 'export.xlsx';
    private $templateFullFilename;
    private $baseDir;
    private $baseFullFilename;
    private $zipFullFilename;
    private $workSheetHandler;
    private $colCount;
    private $rowCount;
    private $numRows = 0;
    private $curCel = 0;
    private $numStrings = 0;
    private $currentRow = array();
    private $isBold = false;

    public function __construct($exportFile, $colCount, $rowCount)
    {
        $this->colCount = $colCount;
        $this->rowCount = $rowCount;

        if (is_null($this->templateFile))
            throw new CHttpException(500,'Invalid template file name \''.$this->templateFile.'\' to export in .xlsx!');

        $this->templateFullFilename = dirname(__FILE__).DIRECTORY_SEPARATOR.'assets'.DIRECTORY_SEPARATOR.$this->templateFile;

        if (!file_exists($this->templateFullFilename))
            throw new CHttpException(500,'Template file name \''.$this->templateFullFilename.'\' not found to export in .xlsx!');

        if (!is_dir($this->exportDir))
            throw new CHttpException(500,'Invalid export directory \''.$this->exportDir.'\' to export in .xlsx!');

        if (!preg_match('/^[a-zA-Z\p{Cyrillic}0-9\-\.\[\]\(\)_]+\.xlsx$/u', $exportFile))
            throw new CHttpException(500, 'Invalid export file name \''.$exportFile.'\' to export in .xlsx!');

        $this->exportFile = $exportFile;

        $this->baseDir = $this->exportDir.DIRECTORY_SEPARATOR.substr($this->exportFile, 0, strlen($this->exportFile) - strlen(pathinfo($this->exportFile, PATHINFO_EXTENSION)) - 1);
        $this->baseFullFilename = $this->exportDir.DIRECTORY_SEPARATOR.$this->exportFile;
        if (file_exists($this->baseFullFilename))
            unlink($this->baseFullFilename);
    }
    
    private function numToLetter($num)
    {
        $f = '';
        do
        {
            $f = chr(64 + $num % 26) . $f;
            $num = floor($num / 26);
        }
        while ($num > 0);
        return $f;
    }

    public function getColCount()
    {
        return $this->colCount;
    }

    public function getRowCount()
    {
        return $this->rowCount;
    }

    public function getBaseFullFileName()
    {
        return $this->baseFullFilename;
    }

    public function getZipFullFileName()
    {
        return $this->zipFullFilename;
    }

    public function openWriter()
    {
        if (is_dir($this->baseDir))
            CFileHelper::removeDirectory($this->baseDir);
        mkdir($this->baseDir);

        exec("unzip $this->templateFullFilename -d \"$this->baseDir\"");

        $this->workSheetHandler = fopen($this->baseDir.'/xl/worksheets/sheet1.xml', 'w+');

        fwrite($this->workSheetHandler, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:'.$this->numToLetter($this->colCount).$this->rowCount.'"/><sheetData>');
    }

    public function resetRow()
    {
        $this->currentRow = array();
    }

    public function openRow($isBold = false)
    {
        $this->isBold = $isBold;
        $this->numRows++;
        $this->currentRow[] = '<row r="'.$this->numRows.'"'.($this->isBold ? ' s="7"' : '').'>';
        $this->curCel = 0;
    }

    public function closeRow()
    {
        $this->currentRow[] = '</row>';
    }

    public function flushRow()
    {
        fwrite($this->workSheetHandler, implode('', $this->currentRow));
        unset($this->currentRow);
    }

    public function appendCellNum($value)
    {
        $this->curCel++;
        $this->currentRow[] = '<c r="'.$this->numToLetter($this->curCel).$this->numRows.'"><v>'.$value.'</v></c>';
    }

    public function appendCellString($value)
    {
        $this->curCel++;
        if (!empty($value)) {
            $value = htmlspecialchars($value, ENT_QUOTES, 'UTF-8');
            $value = preg_replace( '/[\x00-\x13]/', '', $value );
            $this->currentRow[] = '<c r="'.$this->numToLetter($this->curCel).$this->numRows.'" t="inlineStr"'.($this->isBold ? ' s="7"' : '').'><is><t>'.$value.'</t></is></c>';
            $this->numStrings++;
        }
    }

    public function appendCellReal($value)
    {
        return $this->appendCellNum($value);
    }

    public function appendCellDateTime($value)
    {
        $this->curCel++;

        if (empty($value))
            $this->appendCellString('');
        else
        {
            $dt = new DateTime($value);
            $ts = $dt->getTimestamp() + self::ZERO_TIMESTAMP;
            $this->currentRow[] = '<c r="'.$this->numToLetter($this->curCel).$this->numRows.'" s="1"><v>'.$ts/self::SEC_IN_DAY.'</v></c>';
        }
    }

    public function appendCellDate($value)
    {
        $this->curCel++;

        if (empty($value))
            $this->appendCellString('');
        else
        {
            $dt = new DateTime($value);
            $ts = $dt->getTimestamp() + self::ZERO_TIMESTAMP;
            $this->currentRow[] = '<c r="'.$this->numToLetter($this->curCel).$this->numRows.'" s="2"><v>'.floor($ts/self::SEC_IN_DAY).'</v></c>';
        }
    }

    public function appendCellTime($value)
    {
        $this->curCel++;

        if (empty($value))
            $this->appendCellString('');
        else
        {
            $dt = new DateTime($value);
            $ts = $dt->getTimestamp() + self::ZERO_TIMESTAMP;
            $this->currentRow[] = '<c r="'.$this->numToLetter($this->curCel).$this->numRows.'" s="3"><v>'.($ts/self::SEC_IN_DAY - floor($ts/self::SEC_IN_DAY)).'</v></c>';
        }
    }

    public function closeWriter()
    {
        fwrite($this->workSheetHandler, '</sheetData></worksheet>');
        fclose($this->workSheetHandler);
    }

    public function zip()
    {
        $zipfile = '..'.DIRECTORY_SEPARATOR.$this->exportFile;

        $curDir = getcwd();
        chdir($this->baseDir);
        exec("zip -mr \"$zipfile\" *");
        chdir($curDir);

        CFileHelper::removeDirectory($this->baseDir);
        $this->zipFullFilename = $this->exportDir.DIRECTORY_SEPARATOR.$this->exportFile;
    }
}
?>
