<?php
/**
 * Contrôleur Admin Export Clients
 *
 * @author Consogarage
 * @copyright 2025 Consogarage
 * @license Proprietary
 */

if (!defined('_PS_VERSION_')) {
    exit;
}

require_once _PS_MODULE_DIR_ . 'export_clients/classes/ExportClientsGenerator.php';

class AdminExportClientsController extends ModuleAdminController
{
    public function __construct()
    {
        $this->bootstrap = true;
        $this->table = 'customer';
        $this->className = 'Customer';
        $this->lang = false;
        $this->list_no_link = true;

        parent::__construct();
    }

    public function initPageHeaderToolbar()
    {
        parent::initPageHeaderToolbar();
        $this->page_header_toolbar_title = $this->l('Export Clients LD Compta');
    }

    public function initContent()
    {
        $this->content = $this->renderExportForm();
        parent::initContent();
    }

    public function renderList()
    {
        return '';
    }

    private function renderExportForm()
    {
        $fieldsForm = [[
            'form' => [
                'legend' => [
                    'title' => $this->l('Exporter les clients vers LD Compta'),
                    'icon'  => 'icon-download',
                ],
                'description' => $this->l('Sélectionnez une plage de dates pour exporter les clients créés sur cette période, ou laissez les champs vides pour exporter tous les clients.'),
                'input' => [
                    [
                        'type'  => 'date',
                        'label' => $this->l('Date de création — Du'),
                        'name'  => 'export_date_from',
                        'desc'  => $this->l('Laisser vide pour exporter depuis le début'),
                    ],
                    [
                        'type'  => 'date',
                        'label' => $this->l('Date de création — Au'),
                        'name'  => 'export_date_to',
                        'desc'  => $this->l('Laisser vide pour exporter jusqu\'à aujourd\'hui'),
                    ],
                ],
                'buttons' => [
                    [
                        'title' => $this->l('Exporter tous les clients'),
                        'name'  => 'submitExportAll',
                        'type'  => 'submit',
                        'class' => 'btn btn-success pull-left',
                        'icon'  => 'process-icon-download',
                    ],
                ],
                'submit' => [
                    'title' => $this->l('Exporter par période'),
                    'icon'  => 'process-icon-download',
                ],
            ],
        ]];

        $helper = new HelperForm();
        $helper->show_toolbar = false;
        $helper->submit_action = 'submitDateExport';
        $helper->currentIndex = self::$currentIndex;
        $helper->token = $this->token;
        $helper->fields_value = [
            'export_date_from' => Tools::getValue('export_date_from', ''),
            'export_date_to'   => Tools::getValue('export_date_to', ''),
        ];

        return $helper->generateForm($fieldsForm);
    }

    public function postProcess()
    {
        if (Tools::isSubmit('submitExportAll')) {
            $this->runExport(null, null);
            return;
        }
        if (Tools::isSubmit('submitDateExport')) {
            $dateFrom = Tools::getValue('export_date_from');
            $dateTo   = Tools::getValue('export_date_to');

            if ($dateFrom && !Validate::isDate($dateFrom)) {
                $this->errors[] = $this->l('Format de date invalide (Du).');
                return;
            }
            if ($dateTo && !Validate::isDate($dateTo)) {
                $this->errors[] = $this->l('Format de date invalide (Au).');
                return;
            }

            $this->runExport($dateFrom ?: null, $dateTo ?: null);
            return;
        }
        parent::postProcess();
    }

    private function runExport($dateFrom, $dateTo)
    {
        $this->ensureExportMemoryLimit();

        $customers = $this->getCustomers($dateFrom, $dateTo);

        if (empty($customers)) {
            $this->warnings[] = $this->l('Aucun client trouvé pour ces critères.');
            return;
        }

        $pairs = [];
        foreach ($customers as $row) {
            $addressId = $this->getFirstAddressId((int) $row['id_customer']);
            if ($addressId) {
                $pairs[] = ['id_customer' => (int) $row['id_customer'], 'id_address' => $addressId];
            }
        }

        if (empty($pairs)) {
            $this->warnings[] = $this->l('Aucun client avec adresse trouvé.');
            return;
        }

        try {
            $generator = new ExportClientsGenerator();
            $filePath = $generator->generateBatchExport($pairs, $dateFrom, $dateTo);

            if ($filePath && file_exists($filePath)) {
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header('Content-Disposition: attachment; filename="' . basename($filePath) . '"');
                header('Content-Length: ' . filesize($filePath));
                header('Cache-Control: no-cache');
                readfile($filePath);
                exit;
            }
            $this->errors[] = $this->l('Erreur lors de la génération du fichier.');
        } catch (Exception $e) {
            $this->errors[] = $this->l('Erreur : ') . $e->getMessage();
            PrestaShopLogger::addLog('Export Clients: ' . $e->getMessage(), 3);
        }
    }

    private function ensureExportMemoryLimit()
    {
        $currentLimit = ini_get('memory_limit');
        if ($currentLimit === false || $currentLimit === '' || $currentLimit === '-1') {
            return;
        }

        $currentBytes = $this->convertPhpSizeToBytes($currentLimit);
        $targetBytes = 512 * 1024 * 1024;

        if ($currentBytes > 0 && $currentBytes < $targetBytes) {
            @ini_set('memory_limit', '512M');
        }
    }

    private function convertPhpSizeToBytes($value)
    {
        $value = trim((string) $value);
        if ($value === '') {
            return 0;
        }

        $unit = Tools::strtoupper(Tools::substr($value, -1));
        $bytes = (int) $value;

        switch ($unit) {
            case 'G':
                $bytes *= 1024;
            case 'M':
                $bytes *= 1024;
            case 'K':
                $bytes *= 1024;
        }

        return $bytes;
    }

    private function getCustomers($dateFrom, $dateTo)
    {
        $sql = new DbQuery();
        $sql->select('id_customer');
        $sql->from('customer');
        $sql->where('active = 1');
        $sql->where('deleted = 0');

        if ($dateFrom) {
            $sql->where('date_add >= \'' . pSQL($dateFrom . ' 00:00:00') . '\'');
        }
        if ($dateTo) {
            $sql->where('date_add <= \'' . pSQL($dateTo . ' 23:59:59') . '\'');
        }

        $sql->orderBy('date_add ASC');

        return Db::getInstance()->executeS($sql);
    }

    private function getFirstAddressId($customerId)
    {
        $sql = new DbQuery();
        $sql->select('id_address');
        $sql->from('address');
        $sql->where('id_customer = ' . (int) $customerId);
        $sql->where('deleted = 0');
        $sql->orderBy('date_add ASC');

        return (int) Db::getInstance()->getValue($sql);
    }
}
