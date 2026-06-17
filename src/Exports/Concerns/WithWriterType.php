<?php

namespace pxlrbt\FilamentExcel\Exports\Concerns;

use Closure;
use Maatwebsite\Excel\Excel;

trait WithWriterType
{
    /**
     * Holds the user-supplied writer type, which may be a closure resolved at export time.
     *
     * Kept separate from Maatwebsite\Excel\Concerns\Exportable::$writerType (typed ?string),
     * because PHP forbids composed traits declaring the same property with incompatible types.
     */
    protected Closure|string|null $writerTypeCallback = null;

    public function withWriterType(Closure|string|null $writerType = null): static
    {
        $this->writerTypeCallback = $writerType;

        return $this;
    }

    protected function getWriterType(): ?string
    {
        return $this->evaluate($this->writerTypeCallback) ?? Excel::XLSX;
    }

    protected function resolveWriterType(): void
    {
        if ($writerType = data_get($this->formData, 'writer_type')) {
            if ($this->writerTypeCallback instanceof Closure) {
                $writerType = $this->evaluate($this->writerTypeCallback, ['writerType' => $writerType]);
            }

            $this->withWriterType($writerType);
        }
    }

    protected function getDefaultExtension(): string
    {
        return match ($this->getWriterType()) {
            Excel::XLSX => 'xlsx',
            Excel::CSV, Excel::TSV => 'csv',
            Excel::ODS => 'ods',
            Excel::XLS => 'xls',
            Excel::SLK => 'slk',
            Excel::XML => 'xml',
            Excel::GNUMERIC => 'gnumeric',
            Excel::HTML => 'html',
            Excel::MPDF, Excel::DOMPDF, Excel::TCPDF => 'pdf',
            default => 'xlsx',
        };
    }
}
