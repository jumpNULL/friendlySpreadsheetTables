<?php

namespace Crombi\PhpSpreadsheetHelper;

use \PhpOffice\PhpSpreadsheet\Writer\Pdf\Dompdf;

class DomPdfTableExporter extends Dompdf
{
    protected function createExternalWriterInstance()
    {
        return new \Dompdf\Dompdf();
    }
}
