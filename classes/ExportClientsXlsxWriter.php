<?php
/**
 * Écriture d'un fichier XLSX ligne par ligne (une seule feuille), sans dépendance.
 *
 * La feuille est écrite au fil de l'eau dans un fichier temporaire placé à côté
 * du fichier final, puis l'archive XLSX est assemblée avec ZipArchive.
 * Mémoire et disque temporaire constants, quel que soit le nombre de lignes.
 *
 * @author Consogarage
 * @copyright 2025 Consogarage
 * @license Proprietary
 */

if (!defined('_PS_VERSION_')) {
    exit;
}

class ExportClientsXlsxWriter
{
    // Index des styles déclarés dans styles.xml (cellXfs)
    const STYLE_HEADER = 1;
    const STYLE_BODY = 2;
    const STYLE_BODY_TEXT = 3;

    private $filePath;
    private $sheetName;
    private $columns;
    private $textColumns;
    private $sheetTmpPath;
    private $handle;
    private $rowIndex = 0;

    /**
     * @param string $filePath    Chemin du fichier .xlsx à produire
     * @param string $sheetName   Nom de la feuille
     * @param array  $columns     ['A' => largeur, 'B' => largeur, ...] dans l'ordre
     * @param array  $textColumns Colonnes au format Texte (ex. SIRET)
     */
    public function __construct($filePath, $sheetName, array $columns, array $textColumns = [])
    {
        if (!class_exists('ZipArchive')) {
            throw new Exception('Extension PHP zip absente : impossible de générer le fichier XLSX.');
        }

        $this->filePath = $filePath;
        $this->sheetName = $sheetName;
        $this->columns = $columns;
        $this->textColumns = array_flip($textColumns);

        $this->sheetTmpPath = $filePath . '.sheet.tmp';
        $this->handle = @fopen($this->sheetTmpPath, 'wb');
        if (!$this->handle) {
            throw new Exception('Impossible de créer le fichier temporaire de l\'export : ' . $this->sheetTmpPath);
        }

        $this->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . '<cols>');
        $i = 1;
        foreach ($this->columns as $width) {
            $this->write('<col min="' . $i . '" max="' . $i . '" width="' . (float) $width . '" customWidth="1"/>');
            $i++;
        }
        $this->write('</cols><sheetData>');
    }

    /**
     * @param array $labels Libellés dans l'ordre des colonnes
     */
    public function addHeaderRow(array $labels)
    {
        $values = [];
        $letters = array_keys($this->columns);
        foreach (array_values($labels) as $i => $label) {
            $values[$letters[$i]] = $label;
        }

        $this->writeRow($values, true);
    }

    /**
     * @param array $values ['A' => valeur, ...] ; les valeurs vides sont ignorées
     */
    public function addRow(array $values)
    {
        $this->writeRow($values, false);
    }

    public function close()
    {
        $this->write('</sheetData></worksheet>');
        fclose($this->handle);
        $this->handle = null;

        $tmpZip = $this->filePath . '.zip.tmp';
        @unlink($tmpZip);

        try {
            $zip = new ZipArchive();
            if ($zip->open($tmpZip, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
                throw new Exception('Impossible de créer l\'archive XLSX : ' . $tmpZip);
            }

            $zip->addFromString('[Content_Types].xml', $this->contentTypesXml());
            $zip->addFromString('_rels/.rels', $this->rootRelsXml());
            $zip->addFromString('xl/workbook.xml', $this->workbookXml());
            $zip->addFromString('xl/_rels/workbook.xml.rels', $this->workbookRelsXml());
            $zip->addFromString('xl/styles.xml', $this->stylesXml());
            $zip->addFile($this->sheetTmpPath, 'xl/worksheets/sheet1.xml');

            if (!$zip->close()) {
                throw new Exception('Échec de l\'écriture de l\'archive XLSX.');
            }

            if (!@rename($tmpZip, $this->filePath)) {
                throw new Exception('Impossible de finaliser le fichier : ' . $this->filePath);
            }
        } finally {
            @unlink($this->sheetTmpPath);
            @unlink($tmpZip);
        }
    }

    /**
     * Supprime les fichiers temporaires après une erreur.
     */
    public function abort()
    {
        if ($this->handle) {
            @fclose($this->handle);
            $this->handle = null;
        }
        @unlink($this->sheetTmpPath);
        @unlink($this->filePath . '.zip.tmp');
    }

    public function __destruct()
    {
        if ($this->handle) {
            $this->abort();
        }
    }

    private function writeRow(array $values, $isHeader)
    {
        $this->rowIndex++;
        $r = $this->rowIndex;
        $xml = '<row r="' . $r . '">';

        foreach ($this->columns as $col => $width) {
            if (!array_key_exists($col, $values)) {
                continue;
            }
            $value = $values[$col];
            if ($value === '' || $value === null) {
                continue;
            }

            if ($isHeader) {
                $style = self::STYLE_HEADER;
            } elseif (isset($this->textColumns[$col])) {
                $style = self::STYLE_BODY_TEXT;
            } else {
                $style = self::STYLE_BODY;
            }

            $value = (string) $value;

            // Même typage que l'ancien export (DefaultValueBinder de PhpSpreadsheet) :
            // nombre si la valeur est numérique sans zéro initial, texte sinon.
            if (!$isHeader && $style !== self::STYLE_BODY_TEXT && $this->isPlainNumber($value)) {
                $xml .= '<c r="' . $col . $r . '" s="' . $style . '"><v>' . $value . '</v></c>';
                continue;
            }

            $xml .= '<c r="' . $col . $r . '" s="' . $style . '" t="inlineStr"><is><t xml:space="preserve">'
                . $this->escape($value) . '</t></is></c>';
        }

        $this->write($xml . '</row>');
    }

    private function isPlainNumber($value)
    {
        return strlen($value) <= 15 && preg_match('/^-?(0|[1-9][0-9]*)(\.[0-9]+)?$/', $value) === 1;
    }

    private function escape($value)
    {
        // Caractères de contrôle interdits en XML 1.0
        $value = preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F]/u', '', $value);
        if ($value === null) {
            $value = '';
        }

        return htmlspecialchars($value, ENT_QUOTES | ENT_XML1, 'UTF-8');
    }

    private function write($data)
    {
        if (fwrite($this->handle, $data) === false) {
            throw new Exception('Échec d\'écriture dans le fichier temporaire de l\'export (disque plein ?).');
        }
    }

    private function contentTypesXml()
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            . '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            . '<Default Extension="xml" ContentType="application/xml"/>'
            . '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            . '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            . '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            . '</Types>';
    }

    private function rootRelsXml()
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            . '</Relationships>';
    }

    private function workbookXml()
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . '<sheets><sheet name="' . $this->escape($this->sheetName) . '" sheetId="1" r:id="rId1"/></sheets>'
            . '</workbook>';
    }

    private function workbookRelsXml()
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            . '</Relationships>';
    }

    /**
     * Styles : 0 = défaut ; 1 = en-tête (gras, blanc sur bleu 4472C4) ;
     * 2 = corps aligné à gauche ; 3 = corps aligné à gauche au format Texte (@).
     */
    private function stylesXml()
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . '<fonts count="2">'
            . '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
            . '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>'
            . '</fonts>'
            . '<fills count="3">'
            . '<fill><patternFill patternType="none"/></fill>'
            . '<fill><patternFill patternType="gray125"/></fill>'
            . '<fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/><bgColor indexed="64"/></patternFill></fill>'
            . '</fills>'
            . '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            . '<cellXfs count="4">'
            . '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            . '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
            . '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf>'
            . '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left"/></xf>'
            . '</cellXfs>'
            . '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            . '</styleSheet>';
    }
}
