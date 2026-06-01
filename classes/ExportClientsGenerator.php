<?php
/**
 * Générateur d'export clients au format LD Compta
 *
 * @author Consogarage
 * @copyright 2025 Consogarage
 * @license Proprietary
 */

if (!defined('_PS_VERSION_')) {
    exit;
}

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Settings;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

class ExportClientsFileCache implements \Psr\SimpleCache\CacheInterface
{
    private $directory;

    public function __construct($directory)
    {
        $this->directory = rtrim($directory, DIRECTORY_SEPARATOR);

        if (!is_dir($this->directory) && !mkdir($this->directory, 0755, true) && !is_dir($this->directory)) {
            throw new Exception('Impossible de créer le cache disque de l\'export.');
        }
    }

    public function clear()
    {
        foreach (glob($this->directory . DIRECTORY_SEPARATOR . '*.cache') ?: [] as $filePath) {
            if (is_file($filePath)) {
                @unlink($filePath);
            }
        }

        @rmdir($this->directory);

        return true;
    }

    public function delete($key)
    {
        $filePath = $this->getFilePath($key);
        if (is_file($filePath)) {
            @unlink($filePath);
        }

        return true;
    }

    public function deleteMultiple($keys)
    {
        foreach ($keys as $key) {
            $this->delete($key);
        }

        return true;
    }

    public function get($key, $default = null)
    {
        $filePath = $this->getFilePath($key);
        if (!is_file($filePath)) {
            return $default;
        }

        $value = @file_get_contents($filePath);
        if ($value === false) {
            return $default;
        }

        return unserialize($value);
    }

    public function getMultiple($keys, $default = null)
    {
        $results = [];
        foreach ($keys as $key) {
            $results[$key] = $this->get($key, $default);
        }

        return $results;
    }

    public function has($key)
    {
        return is_file($this->getFilePath($key));
    }

    public function set($key, $value, $ttl = null)
    {
        $filePath = $this->getFilePath($key);

        return file_put_contents($filePath, serialize($value), LOCK_EX) !== false;
    }

    public function setMultiple($values, $ttl = null)
    {
        foreach ($values as $key => $value) {
            if (!$this->set($key, $value, $ttl)) {
                return false;
            }
        }

        return true;
    }

    private function getFilePath($key)
    {
        return $this->directory . DIRECTORY_SEPARATOR . sha1((string) $key) . '.cache';
    }
}

class ExportClientsGenerator
{
    private $cellCache;

    /**
     * Colonnes LD Compta (A–Y) avec leur largeur d'affichage
     */
    private static $columns = [
        'A'  => ['code' => 'TYPE', 'width' => 6],
        'B'  => ['code' => 'NOCL', 'width' => 10],
        'C'  => ['code' => 'NMCM', 'width' => 22],
        'D'  => ['code' => 'LIBI', 'width' => 42],
        'E'  => ['code' => 'TITR', 'width' => 6],
        'F'  => ['code' => 'RSSO', 'width' => 42],
        'G'  => ['code' => 'CAD1', 'width' => 42],
        'H'  => ['code' => 'CADR', 'width' => 42],
        'I'  => ['code' => 'NVOI', 'width' => 42],
        'J'  => ['code' => 'COPO', 'width' => 7],
        'K'  => ['code' => 'BUDI', 'width' => 42],
        'L'  => ['code' => 'CPAY', 'width' => 5],
        'M'  => ['code' => 'PAYS', 'width' => 32],
        'N'  => ['code' => 'NTEL', 'width' => 22],
        'O'  => ['code' => 'TLEX', 'width' => 22],
        'P'  => ['code' => 'TLPO', 'width' => 22],
        'Q'  => ['code' => 'TLCY', 'width' => 22],
        'R'  => ['code' => 'NINT', 'width' => 42],
        'S'  => ['code' => 'COMM', 'width' => 15],
        'T'  => ['code' => 'SIRE', 'width' => 16],
        'U'  => ['code' => 'RIBP', 'width' => 5],
        'V'  => ['code' => 'DOBQ', 'width' => 27],
        'W'  => ['code' => 'IBBQ', 'width' => 6],
        'X'  => ['code' => 'COBQ', 'width' => 7],
        'Y'  => ['code' => 'GUBQ', 'width' => 7],
        'Z'  => ['code' => 'CPBQ', 'width' => 13],
        'AA' => ['code' => 'CLBQ', 'width' => 11],
        'AB' => ['code' => 'BIBQ', 'width' => 13],
        'AC' => ['code' => 'MOPM', 'width' => 5],
        'AD' => ['code' => 'DJPM', 'width' => 6],
        'AE' => ['code' => 'DMPM', 'width' => 6],
        'AF' => ['code' => 'REFM', 'width' => 5],
        'AG' => ['code' => 'SLVA', 'width' => 5],
        'AH' => ['code' => 'PLCR', 'width' => 5],
        'AI' => ['code' => 'ECFI', 'width' => 5],
        'AJ' => ['code' => 'CREP', 'width' => 5],
        'AK' => ['code' => 'NREP', 'width' => 5],
        'AL' => ['code' => 'TREP', 'width' => 5],
        'AM' => ['code' => 'MREP', 'width' => 5],
        'AN' => ['code' => 'GRRE', 'width' => 5],
        'AO' => ['code' => 'LTTA', 'width' => 5],
        'AP' => ['code' => 'CACT', 'width' => 5],
        'AQ' => ['code' => 'CODV', 'width' => 5],
        'AR' => ['code' => 'GRTR', 'width' => 5],
        'AS' => ['code' => 'NOFP', 'width' => 5],
        'AT' => ['code' => 'BQAF', 'width' => 5],
        'AU' => ['code' => 'BONP', 'width' => 5],
        'AV' => ['code' => 'CESC', 'width' => 5],
        'AW' => ['code' => 'EWEB', 'width' => 42],
        'AX' => ['code' => 'EMAI', 'width' => 42],
        'AY' => ['code' => 'NAPE', 'width' => 5],
        'AZ' => ['code' => 'NCCE', 'width' => 5],
        'BA' => ['code' => 'SUSP', 'width' => 5],
        'BB' => ['code' => 'CLGE', 'width' => 5],
        'BC' => ['code' => 'NEDI', 'width' => 5],
        'BD' => ['code' => 'CFAM', 'width' => 5],
        'BE' => ['code' => 'CGRP', 'width' => 5],
        'BF' => ['code' => 'RACI', 'width' => 5],
        'BG' => ['code' => 'CZC1', 'width' => 5],
        'BH' => ['code' => 'CZC2', 'width' => 5],
        'BI' => ['code' => 'CZC3', 'width' => 5],
        'BJ' => ['code' => 'CZC4', 'width' => 5],
        'BK' => ['code' => 'CZC5', 'width' => 5],
        'BL' => ['code' => 'CZC6', 'width' => 5],
        'BM' => ['code' => 'CZC7', 'width' => 5],
        'BN' => ['code' => 'CZC8', 'width' => 5],
        'BO' => ['code' => 'DTC1', 'width' => 5],
        'BP' => ['code' => 'DTC2', 'width' => 5],
        'BQ' => ['code' => 'DTC3', 'width' => 5],
        'BR' => ['code' => 'DTC4', 'width' => 5],
        'BS' => ['code' => 'ZNC1', 'width' => 5],
        'BT' => ['code' => 'ZNC2', 'width' => 5],
        'BU' => ['code' => 'ZNC3', 'width' => 5],
        'BV' => ['code' => 'RTOB', 'width' => 5],
    ];

    /**
     * Générer l'export LD Compta pour un seul client
     *
     * @param int $customerId
     * @param int $addressId
     * @return string Chemin du fichier généré
     */
    public function generateExport($customerId, $addressId)
    {
        $customer = new Customer((int) $customerId);
        $address = new Address((int) $addressId);

        if (!Validate::isLoadedObject($customer) || !Validate::isLoadedObject($address)) {
            throw new Exception('Client ou adresse introuvable');
        }

        $spreadsheet = $this->createSpreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $this->addHeaderRow($sheet);
        $this->addCustomerRow($sheet, 2, $customer, $address);
        $this->applyGeneralStyling($sheet, 2);

        $filename = $this->generateFilename($customer);
        $filePath = $this->getExportDir() . $filename;

        $this->saveSpreadsheet($spreadsheet, $filePath);

        return $filePath;
    }

    /**
     * Générer l'export LD Compta pour une liste de clients (par plage de dates)
     *
     * @param array  $pairs     [['id_customer' => x, 'id_address' => y], ...]
     * @param string $dateFrom  YYYY-MM-DD
     * @param string $dateTo    YYYY-MM-DD
     * @return string Chemin du fichier généré
     */
    public function generateBatchExport(array $pairs, $dateFrom = null, $dateTo = null)
    {
        $spreadsheet = $this->createSpreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $this->addHeaderRow($sheet);

        $row = 2;
        foreach ($pairs as $pair) {
            $customer = new Customer((int) $pair['id_customer']);
            $address = new Address((int) $pair['id_address']);
            if (Validate::isLoadedObject($customer) && Validate::isLoadedObject($address)) {
                $this->addCustomerRow($sheet, $row, $customer, $address);
                $row++;
            }
        }

        $this->applyGeneralStyling($sheet, max(2, $row - 1));

        $suffix = ($dateFrom && $dateTo)
            ? '_' . str_replace('-', '', $dateFrom) . '_' . str_replace('-', '', $dateTo)
            : '_' . date('Ymd');
        $filePath = $this->getExportDir() . 'export_clients' . $suffix . '.xlsx';

        $this->saveSpreadsheet($spreadsheet, $filePath);

        return $filePath;
    }

    private function createSpreadsheet()
    {
        $this->initializeCellCache();

        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->setTitle('Tiers');

        return $spreadsheet;
    }

    private function initializeCellCache()
    {
        if ($this->cellCache instanceof ExportClientsFileCache) {
            $this->cellCache->clear();
        }

        $cacheDir = rtrim(sys_get_temp_dir(), DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . 'export_clients_cache_' . uniqid('', true);
        $this->cellCache = new ExportClientsFileCache($cacheDir);
        Settings::setCache($this->cellCache);
    }

    private function saveSpreadsheet(Spreadsheet $spreadsheet, $filePath)
    {
        $writer = new Xlsx($spreadsheet);

        try {
            $writer->save($filePath);
        } finally {
            $spreadsheet->disconnectWorksheets();
            unset($writer, $spreadsheet);

            if ($this->cellCache instanceof ExportClientsFileCache) {
                $this->cellCache->clear();
                $this->cellCache = null;
            }
        }
    }

    private function addHeaderRow($sheet)
    {
        foreach (self::$columns as $col => $def) {
            $sheet->setCellValue($col . '1', $def['code']);
        }

        $lastCol = array_key_last(self::$columns);
        $range = 'A1:' . $lastCol . '1';

        $sheet->getStyle($range)->getFont()->setBold(true);
        $sheet->getStyle($range)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB('4472C4');
        $sheet->getStyle($range)->getFont()->getColor()->setRGB('FFFFFF');
    }

    private function addCustomerRow($sheet, $row, $customer, $address)
    {
        $country = new Country((int) $address->id_country);
        $isoCode = Tools::strtoupper((string) $country->iso_code);
        $countryName = (string) Country::getNameById(
            (int) Context::getContext()->language->id,
            (int) $address->id_country
        );

        $nomAbrege = Tools::strtoupper(
            $address->company
                ? Tools::substr((string) $address->company, 0, 20)
                : Tools::substr($customer->lastname . ' ' . $customer->firstname, 0, 20)
        );

        $libelleInterne = Tools::strtoupper(
            $address->company
                ? Tools::substr((string) $address->company, 0, 40)
                : Tools::substr($customer->lastname . ' ' . $customer->firstname, 0, 40)
        );

        $siret = !empty($customer->siret)
            ? Tools::substr(preg_replace('/\s+/', '', (string) $customer->siret), 0, 14)
            : '';

        $raisonSociale = !empty($address->company)
            ? Tools::strtoupper(Tools::substr((string) $address->company, 0, 40))
            : '';

        $mopm = in_array(4, $customer->getGroups()) ? 'TD' : 'CB';

        $values = [
            'A'  => 'C',
            'B'  => str_pad((string) (int) $customer->id, 8, '0', STR_PAD_LEFT),
            'C'  => $nomAbrege,
            'D'  => $libelleInterne,
            'E'  => '',
            'F'  => $raisonSociale,
            'G' => Tools::substr(trim((string) $address->address1), 0, 40),
            'H' => Tools::substr(trim((string) $address->address2), 0, 40),
            'I' => '',
            'J' => Tools::substr((string) $address->postcode, 0, 5),
            'K' => Tools::substr((string) $address->city, 0, 40),
            'L' => Tools::substr($isoCode, 0, 2),
            'M' => Tools::substr($countryName, 0, 30),
            'N' => Tools::substr((string) $address->phone, 0, 20),
            'O' => '',
            'P' => Tools::substr((string) $address->phone_mobile, 0, 20),
            'Q' => '',
            'R' => Tools::substr($customer->firstname . ' ' . $customer->lastname, 0, 40),
            'S' => '',
            'T' => $siret,
            'U'  => '',
            'V'  => '',
            'W'  => '',
            'X'  => '',
            'Y'  => '',
            'Z'  => '',
            'AA' => '',
            'AB' => '',
            'AC' => $mopm,
            'AD' => $mopm === 'TD' ? '30' : '0',
            'AE' => '',
            'AF' => '',
            'AG' => 'O',
            'AH' => '',
            'AI' => '',
            'AJ' => '',
            'AK' => '',
            'AL' => '',
            'AM' => '',
            'AN' => '1',
            'AO' => '',
            'AP' => '',
            'AQ' => '',
            'AR' => '',
            'AS' => '',
            'AT' => '',
            'AU' => '',
            'AV' => '',
            'AW' => '',
            'AX' => '',
            'AY' => '',
            'AZ' => '',
            'BA' => '',
            'BB' => '',
            'BC' => '',
            'BD' => '',
            'BE' => '',
            'BF' => '',
            'BG' => '',
            'BH' => '',
            'BI' => '',
            'BJ' => '',
            'BK' => '',
            'BL' => '',
            'BM' => '',
            'BN' => '',
            'BO' => '',
            'BP' => '',
            'BQ' => '',
            'BR' => '',
            'BS' => '',
            'BT' => '',
            'BU' => '',
            'BV' => '',
        ];

        foreach ($values as $col => $value) {
            if ($value === '' || $value === null) {
                continue;
            }

            if ($col === 'T') {
                $sheet->setCellValueExplicit($col . $row, (string) $value, DataType::TYPE_STRING);
                continue;
            }

            $sheet->setCellValue($col . $row, $value);
        }
    }

    private function applyGeneralStyling($sheet, $lastRow)
    {
        foreach (self::$columns as $col => $def) {
            $sheet->getColumnDimension($col)->setWidth($def['width']);
        }

        $sheet->getStyle('T:T')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_TEXT);

        if ($lastRow >= 2) {
            $lastCol = array_key_last(self::$columns);
            $sheet->getStyle('A2:' . $lastCol . $lastRow)
                ->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_LEFT);
        }
    }

    private function getExportDir()
    {
        $dir = _PS_MODULE_DIR_ . 'export_clients/exports/';
        if (!file_exists($dir)) {
            mkdir($dir, 0755, true);
        }
        return $dir;
    }

    private function generateFilename($customer)
    {
        $date = date('Y-m-d');
        $name = Tools::strtoupper(Tools::replaceAccentedChars((string) $customer->lastname));
        return 'export_client_' . (int) $customer->id . '_' . $name . '_' . $date . '.xlsx';
    }
}
