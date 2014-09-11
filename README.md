AlxdExportXLSX
==============

Class for export data to Microsoft Excel in format XLSX

Simple example (from class):

```php
public function exportXLSX(&$filename)
{
    $this->_provider = new CArrayDataProvider(/*query*/);

    Yii::import('ext.AlxdExportXLSX.AlxdExportXLSX');
    $export = new AlxdExportXLSX($filename, count($this->_attributes), $this->_provider->getTotalItemCount() + 1);

    $export->openWriter();
    $export->resetRow();
    $export->openRow(true);
    foreach ($this->_attributes as $code => $format)
        $export->appendCellString($this->_objectref->getAttributeLabel($code));
    $export->closeRow();
    $export->flushRow();

    $rows = new CDataProviderIterator($this->_provider, 100);
    foreach ($rows as $row)
    {
        $export->resetRow();
        $export->openRow();

        foreach ($this->_attributes as $code => $format)
        {
            switch ($format->type)
            {
                case 'Num':
                    $export->appendCellNum($row[$code]);
                /*other types*/
                default:
                    $export->appendCellString('');					
            }
        }

        $export->closeRow();
        $export->flushRow();
    }
    $export->closeWriter();
    $export->zip();

    $filename = $export->getZipFullFileName();
}
```
