<?php
/**
 * Générateur d'export clients au format XLSM
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
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class ExportClientsGenerator
{
    /**
     * Générer l'export XLSM pour un client
     *
     * @param int $customerId
     * @param int $orderId
     * @return string|false Le chemin du fichier généré ou false en cas d'erreur
     */
    public function generateExport($customerId, $orderId)
    {
        $customer = new Customer($customerId);
        $order = new Order($orderId);

        if (!Validate::isLoadedObject($customer) || !Validate::isLoadedObject($order)) {
            throw new Exception('Client ou commande introuvable');
        }

        // Récupérer les adresses
        $addressDelivery = new Address($order->id_address_delivery);
        $addressInvoice = new Address($order->id_address_invoice);

        // Créer le spreadsheet
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Titre du document
        $sheet->setTitle('Export Client');

        // Section Coordonnées
        $this->addCoordinatesSection($sheet, $customer, $addressInvoice);

        // Section Détail
        $this->addDetailSection($sheet, $customer, $addressInvoice);

        // Section Paiement
        $this->addPaymentSection($sheet, $order, $customer);

        // Style général
        $this->applyGeneralStyling($sheet);

        // Générer le nom du fichier
        $filename = $this->generateFilename($customer);
        $exportDir = _PS_MODULE_DIR_ . 'export_clients/exports/';

        if (!file_exists($exportDir)) {
            mkdir($exportDir, 0755, true);
        }

        $filePath = $exportDir . $filename;

        // Sauvegarder le fichier
        $writer = new Xlsx($spreadsheet);
        $writer->save($filePath);

        return $filePath;
    }

    /**
     * Ajouter la section Coordonnées
     */
    private function addCoordinatesSection($sheet, $customer, $address)
    {
        $row = 1;

        // Titre de section
        $sheet->setCellValue("A{$row}", 'COORDONNÉES');
        $this->applySectionTitleStyle($sheet, "A{$row}");
        $row += 2;

        // Titre (Civilité)
        $gender = new Gender($customer->id_gender, Context::getContext()->language->id);
        $sheet->setCellValue("A{$row}", 'Titre');
        $sheet->setCellValue("B{$row}", $gender->name ? $gender->name : '');
        $row++;

        // Raison sociale
        $sheet->setCellValue("A{$row}", 'Raison sociale');
        $sheet->setCellValue("B{$row}", $address->company ? $address->company : '');
        $row++;

        // Adresse
        $sheet->setCellValue("A{$row}", 'Adresse');
        $fullAddress = trim($address->address1 . ' ' . $address->address2);
        $sheet->setCellValue("B{$row}", $fullAddress);
        $row++;

        // Code postal
        $sheet->setCellValue("A{$row}", 'Code postal');
        $sheet->setCellValue("B{$row}", $address->postcode);
        $row++;

        // Ville
        $sheet->setCellValue("A{$row}", 'Ville');
        $sheet->setCellValue("B{$row}", $address->city);
        $row++;

        // Pays
        $country = new Country($address->id_country, Context::getContext()->language->id);
        $sheet->setCellValue("A{$row}", 'Pays');
        $sheet->setCellValue("B{$row}", $country->name);
        $row++;

        // Téléphone
        $sheet->setCellValue("A{$row}", 'Téléphone');
        $sheet->setCellValue("B{$row}", $address->phone ? $address->phone : '');
        $row++;

        // Bureau (vide par défaut)
        $sheet->setCellValue("A{$row}", 'Bureau');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        // Télécopie (vide par défaut)
        $sheet->setCellValue("A{$row}", 'Télécopie');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        // Portable
        $sheet->setCellValue("A{$row}", 'Portable');
        $sheet->setCellValue("B{$row}", $address->phone_mobile ? $address->phone_mobile : '');
        $row++;

        // Adresse électronique (email)
        $sheet->setCellValue("A{$row}", 'Adresse électronique');
        $sheet->setCellValue("B{$row}", $customer->email);
        $row++;

        // Site web (vide par défaut)
        $sheet->setCellValue("A{$row}", 'Site web');
        $sheet->setCellValue("B{$row}", $address->website ? $address->website : '');
        $row++;

        return $row;
    }

    /**
     * Ajouter la section Détail
     */
    private function addDetailSection($sheet, $customer, $address)
    {
        $row = 15;

        // Titre de section
        $sheet->setCellValue("A{$row}", 'DÉTAIL');
        $this->applySectionTitleStyle($sheet, "A{$row}");
        $row += 2;

        // Classification
        $sheet->setCellValue("A{$row}", 'Classification');
        $this->applySubSectionStyle($sheet, "A{$row}");
        $row++;

        $sheet->setCellValue("A{$row}", 'Code groupe');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        $sheet->setCellValue("A{$row}", 'Code famille');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        $sheet->setCellValue("A{$row}", 'Collectif');
        $sheet->setCellValue("B{$row}", '');
        $row += 2;

        // Renseignements complémentaires
        $sheet->setCellValue("A{$row}", 'Renseignements complémentaires');
        $this->applySubSectionStyle($sheet, "A{$row}");
        $row++;

        $sheet->setCellValue("A{$row}", 'Numéro de SIRET');
        $sheet->setCellValue("B{$row}", $address->dni ? $address->dni : '');
        $row++;

        $sheet->setCellValue("A{$row}", 'Code NAF');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        $sheet->setCellValue("A{$row}", 'N° TVA intracommunautaire');
        $sheet->setCellValue("B{$row}", $address->vat_number ? $address->vat_number : '');
        $row++;

        $sheet->setCellValue("A{$row}", 'N° EDI');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        $sheet->setCellValue("A{$row}", 'Nom interlocuteur');
        $sheet->setCellValue("B{$row}", $customer->firstname . ' ' . $customer->lastname);
        $row++;

        $sheet->setCellValue("A{$row}", 'Commentaires');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        $sheet->setCellValue("A{$row}", 'Code TVA');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        return $row;
    }

    /**
     * Ajouter la section Paiement
     */
    private function addPaymentSection($sheet, $order, $customer)
    {
        $row = 35;

        // Titre de section
        $sheet->setCellValue("A{$row}", 'PAIEMENT');
        $this->applySectionTitleStyle($sheet, "A{$row}");
        $row += 2;

        // Conditions de paiement
        $sheet->setCellValue("A{$row}", 'Conditions de paiement');
        $this->applySubSectionStyle($sheet, "A{$row}");
        $row++;

        // Mode de paiement
        $sheet->setCellValue("A{$row}", 'Mode de paiement');
        $sheet->setCellValue("B{$row}", $order->payment);
        $row++;

        // Délai de paiement
        $sheet->setCellValue("A{$row}", 'Délai de paiement');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        // Décalage
        $sheet->setCellValue("A{$row}", 'Décalage');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        // Type
        $sheet->setCellValue("A{$row}", 'Type');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        // Banque d'affectation
        $sheet->setCellValue("A{$row}", 'Banque d\'affectation');
        $sheet->setCellValue("B{$row}", '');
        $row += 2;

        // Coordonnées bancaires
        $sheet->setCellValue("A{$row}", 'Coordonnées bancaires');
        $this->applySubSectionStyle($sheet, "A{$row}");
        $row++;

        // Domiciliation
        $sheet->setCellValue("A{$row}", 'Domiciliation');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        // IBAN
        $sheet->setCellValue("A{$row}", 'IBAN');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        // RIB
        $sheet->setCellValue("A{$row}", 'RIB');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        // Code BIC
        $sheet->setCellValue("A{$row}", 'Code BIC');
        $sheet->setCellValue("B{$row}", '');
        $row++;

        return $row;
    }

    /**
     * Appliquer le style pour les titres de section
     */
    private function applySectionTitleStyle($sheet, $cell)
    {
        $sheet->getStyle($cell)->getFont()->setBold(true)->setSize(14);
        $sheet->getStyle($cell)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB('4472C4');
        $sheet->getStyle($cell)->getFont()->getColor()->setRGB('FFFFFF');
    }

    /**
     * Appliquer le style pour les sous-sections
     */
    private function applySubSectionStyle($sheet, $cell)
    {
        $sheet->getStyle($cell)->getFont()->setBold(true)->setSize(11);
        $sheet->getStyle($cell)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB('D9E1F2');
    }

    /**
     * Appliquer le style général
     */
    private function applyGeneralStyling($sheet)
    {
        // Largeur des colonnes
        $sheet->getColumnDimension('A')->setWidth(30);
        $sheet->getColumnDimension('B')->setWidth(40);

        // Style des labels (colonne A)
        $sheet->getStyle('A:A')->getFont()->setBold(true);
        $sheet->getStyle('A:A')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Style des valeurs (colonne B)
        $sheet->getStyle('B:B')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
    }

    /**
     * Générer le nom du fichier
     */
    private function generateFilename($customer)
    {
        $date = date('Y-m-d');
        $customerName = Tools::strtoupper(Tools::replaceAccentedChars($customer->lastname));
        $customerId = $customer->id;

        return "export_client_{$customerId}_{$customerName}_{$date}.xlsx";
    }
}
