<?php
namespace hhashimoto\excel;

class Book {
    // string
    private $name;

    // Sheets
    private $sheets;

    // SharedStrings
    private $sharedStrings;

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
        $this->sheets = null;
        $this->sharedStrings = null;

        try {
            // SharedString
            $ss = $zip->getFromName('xl/sharedStrings.xml');
            $xml = new \SimpleXMLElement($ss);
            $ss = null;

            $sharedStrings = new SharedStrings;
            foreach ($xml->si as $si) {
                if (isset($si->r)) {
                    $t = '';
                    foreach ($si->r as $r) {
                        $t .= $r->t;
                    }
                    $sharedStrings->add($t);
                } else {
                    $sharedStrings->add($si->t);
                }
            }
            $this->sharedStrings = $sharedStrings;

            // WorkBook
            $workbook = $zip->getFromName('xl/workbook.xml');
            $zip->close();
            $zip = null;

            $xml = new \SimpleXMLElement($workbook);
            $workbook = null;

            Sheet::belongsTo($this);
            $sheets = new Sheets;
            foreach ($xml->sheets->sheet as $tag => $sheet) {
                $r = $sheet->attributes('r', true);
                $s = new Sheet($sheet['name'], $sheet['sheetId'], $r['id']);
                $sheets->add($s);
            }
            $this->sheets = $sheets;

        } catch (\Exception $e) {
            echo __FILE__ . ' : ' . __LINE__ . ' ' . $e . PHP_EOL;

            $this->name = null;
            $this->sheets = null;
            $this->sharedStrings = null;
            Sheet::belongsTo(null);
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

    /**
     * @param int $index integer or numeric string
     * @return string
     */
    public function getSharedString($index) {
        return $this->sharedStrings->get((int)$index);
    }

    /**
     * @return array
     */
    public function getSheetNames() {
        return $this->sheets->sheetNames();
    }
}
