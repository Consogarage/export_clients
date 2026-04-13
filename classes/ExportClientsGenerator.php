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
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class ExportClientsGenerator
{
    /**
     * Colonnes LD Compta (A–Y) avec leur largeur d'affichage
     */
    private static $columns = [
        'A' => ['code' => 'TYPE', 'width' => 6],
        'B' => ['code' => 'NOCL', 'width' => 10],
        'C' => ['code' => 'NMCM', 'width' => 22],
        'D' => ['code' => 'LIBI', 'width' => 42],
        'E' => ['code' => 'TITR', 'width' => 6],
        'F' => ['code' => 'RSSO', 'width' => 42],
        'G' => ['code' => 'CAD1', 'width' => 42],
        'H' => ['code' => 'CADR', 'width' => 42],
        'I' => ['code' => 'NVOI', 'width' => 42],
        'J' => ['code' => 'COPO', 'width' => 7],
        'K' => ['code' => 'BUDI', 'width' => 42],
        'L' => ['code' => 'CPAY', 'width' => 5],
        'M' => ['code' => 'PAYS', 'width' => 32],
        'N' => ['code' => 'NTEL', 'width' => 22],
        'O' => ['code' => 'TLEX', 'width' => 22],
        'P' => ['code' => 'TLPO', 'width' => 22],
        'Q' => ['code' => 'TLCY', 'width' => 22],
        'R' => ['code' => 'NINT', 'width' => 42],
        'S' => ['code' => 'COMM', 'width' => 15],
        'T' => ['code' => 'SIRE', 'width' => 16],
        'U' => ['code' => 'RIBP', 'width' => 5],
        'V' => ['code' => 'DOBQ', 'width' => 27],
        'W' => ['code' => 'IBBQ', 'width' => 6],
        'X' => ['code' => 'COBQ', 'width' => 7],
        'Y' => ['code' => 'GUBQ', 'width' => 7],        'Z'  => ['code' => 'CPBQ', 'width' => 13],
        'AA' => ['code' => 'CLBQ', 'width' => 11],
        'AB' => ['code' => 'BIBQ', 'width' => 13],
        'AC' => ['code' => 'MOPM', 'width' => 5],
        'AD' => ['code' => 'DJPM', 'width' => 6],
        'AE' => ['code' => 'DMPM', 'width' => 6],
        'AF' => ['code' => 'REFM', 'width' => 5],    ];

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

        (new Xlsx($spreadsheet))->save($filePath);

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

        (new Xlsx($spreadsheet))->save($filePath);

        return $filePath;
    }

    private function createSpreadsheet()
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->setTitle('Tiers');
        return $spreadsheet;
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

        $siret = $address->dni
            ? Tools::substr(preg_replace('/\s+/', '', (string) $address->dni), 0, 14)
            : '';

        $values = [
            'A' => 'C',
            'B' => str_pad((string) (int) $customer->id, 8, '0', STR_PAD_LEFT),
            'C' => $nomAbrege,
            'D' => $libelleInterne,
            'E' => '',
            'F' => '',
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
            'AC' => 'TD',
            'AD' => '30',
            'AE' => '0',
            'AF' => 'M',
        ];

        foreach ($values as $col => $value) {
            $sheet->setCellValue($col . $row, $value);
        }
    }

    private function applyGeneralStyling($sheet, $lastRow)
    {
        foreach (self::$columns as $col => $def) {
            $sheet->getColumnDimension($col)->setWidth($def['width']);
        }

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
