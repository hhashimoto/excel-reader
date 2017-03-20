<?php

class Book {
    // string
    private $name;

    // Sheet
    private $sheets;

    /**
     * Load Excel-Book named $name
     *
     * @param string $name
     */
    public function load($name) {
        if (! file_exists($name)) {
            throw new \Exception("'{$name}' not found!");
        }

        $zip = new \ZipArchive;
        if (! $zip->open($name)) {
            throw new \Exception("'{$name}' could not open!");
        }

        $this->name = $name;

        try {
            $workbook = $zip->getFromName('xl/workbook.xml');
            $zip->close();
            $zip = null;

            $xml = new SimpleXMLElement($workbook);
            $workbook = null;

            $sheets = new Sheets;
            foreach ($xml->sheets->sheet as $tag => $sheet) {
                $r = $sheet->attributes('r', true);
                $s = new Sheet($sheet['name'], $sheet['sheetId'], $r['id']);
                $sheets->add($s);
            }
            $this->sheets = $sheets;

        } catch (\Exception $e) {
            echo __FILE__ . ' : ' . __LINE__ . ' ' . $e . PHP_EOL;
        } finally {
            if ($zip) {
                $zip->close();
                $zip = null;
            }
        }
    }

    private function sheets() {
        if (! $this->sheets) {
            throw new \Exception('No book has been loaded!');
        }
        return $this->sheets;
    }

    /**
     * Return number of sheets in the loaded book
     *
     * @return int
     */
    public function numberOfSheets() {
        return $this->sheets()->count();
    }

    /**
     * Return sheet named $name
     *
     * @param string $name
     * @return (Sheet | null)
     */
    public function getSheetByName($name) {
        return $this->sheets()->find($name);
    }

    /**
     * Return name of loaded book
     *
     * @return string
     */
    public function name() {
        return $this->name;
    }
}
