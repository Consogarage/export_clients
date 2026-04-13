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
        $this->addRowAction('export');
        $this->addRowAction('view');

        $this->bulk_actions = [
            'export' => [
                'text' => $this->l('Export sélection'),
                'icon' => 'icon-download',
            ],
        ];

        $this->fields_list = [
            'id_customer' => [
                'title' => $this->l('ID'),
                'align' => 'center',
                'class' => 'fixed-width-xs',
            ],
            'company' => [
                'title' => $this->l('Raison sociale'),
            ],
            'lastname' => [
                'title' => $this->l('Nom'),
            ],
            'firstname' => [
                'title' => $this->l('Prénom'),
            ],
            'email' => [
                'title' => $this->l('Email'),
            ],
            'total_orders' => [
                'title' => $this->l('Nombre de commandes'),
                'align' => 'center',
                'search' => false,
                'orderby' => false,
            ],
            'date_add' => [
                'title' => $this->l('Date d\'inscription'),
                'type' => 'date',
                'align' => 'right',
            ],
        ];

        parent::__construct();
    }

    /**
     * Récupérer la liste des clients avec le nombre de commandes
     */
    public function getList($id_lang, $orderBy = null, $orderWay = null, $start = 0, $limit = null, $id_lang_shop = null)
    {
        parent::getList($id_lang, $orderBy, $orderWay, $start, $limit, $id_lang_shop);

        // Ajouter le nombre de commandes pour chaque client
        if (!empty($this->_list)) {
            foreach ($this->_list as &$row) {
                $row['total_orders'] = $this->getCustomerOrderCount($row['id_customer']);
            }
        }
    }

    /**
     * Obtenir le nombre de commandes d'un client
     */
    private function getCustomerOrderCount($customerId)
    {
        $sql = new DbQuery();
        $sql->select('COUNT(*)');
        $sql->from('orders');
        $sql->where('id_customer = ' . (int) $customerId);
        $sql->where('current_state != ' . (int) Configuration::get('PS_OS_CANCELED'));
        $sql->where('current_state != ' . (int) Configuration::get('PS_OS_ERROR'));

        return (int) Db::getInstance()->getValue($sql);
    }

    /**
     * Action d'export pour un client
     */
    public function displayExportLink($token = null, $id, $name = null)
    {
        $this->context->smarty->assign([
            'href' => self::$currentIndex .
                '&' . $this->identifier . '=' . $id .
                '&export' . $this->table . '&token=' . ($token != null ? $token : $this->token),
            'action' => $this->l('Exporter'),
            'icon' => 'icon-download',
        ]);

        return $this->context->smarty->fetch('controllers/admin/_partials/action_button.tpl');
    }

    /**
     * Process de l'action d'export
     */
    public function processExport()
    {
        $customerId = (int) Tools::getValue('id_customer');

        if (!$customerId) {
            $this->errors[] = $this->l('ID client invalide');
            return;
        }

        try {
            $generator = new ExportClientsGenerator();

            // Récupérer la première commande du client
            $orderId = $this->getFirstOrderId($customerId);

            if (!$orderId) {
                $this->errors[] = $this->l('Aucune commande trouvée pour ce client');
                return;
            }

            $filePath = $generator->generateExport($customerId, $orderId);

            if ($filePath && file_exists($filePath)) {
                $this->confirmations[] = $this->l('Export généré avec succès');

                // Télécharger le fichier
                header('Content-Type: application/vnd.ms-excel.sheet.macroEnabled.12');
                header('Content-Disposition: attachment; filename="' . basename($filePath) . '"');
                header('Content-Length: ' . filesize($filePath));
                readfile($filePath);
                exit;
            } else {
                $this->errors[] = $this->l('Erreur lors de la génération du fichier');
            }
        } catch (Exception $e) {
            $this->errors[] = $this->l('Erreur: ') . $e->getMessage();
            PrestaShopLogger::addLog(
                'Export Clients: Erreur - ' . $e->getMessage(),
                3,
                null,
                'Customer',
                $customerId
            );
        }
    }

    /**
     * Process de l'action d'export en masse
     */
    public function processBulkExport()
    {
        $customerIds = Tools::getValue($this->table . 'Box');

        if (!is_array($customerIds) || empty($customerIds)) {
            $this->errors[] = $this->l('Aucun client sélectionné');
            return;
        }

        $exportedCount = 0;
        $errorCount = 0;

        foreach ($customerIds as $customerId) {
            try {
                $generator = new ExportClientsGenerator();
                $orderId = $this->getFirstOrderId((int) $customerId);

                if ($orderId) {
                    $generator->generateExport((int) $customerId, $orderId);
                    $exportedCount++;
                } else {
                    $errorCount++;
                }
            } catch (Exception $e) {
                $errorCount++;
                PrestaShopLogger::addLog(
                    'Export Clients: Erreur export bulk - ' . $e->getMessage(),
                    3,
                    null,
                    'Customer',
                    (int) $customerId
                );
            }
        }

        if ($exportedCount > 0) {
            $this->confirmations[] = sprintf($this->l('%d client(s) exporté(s) avec succès'), $exportedCount);
        }

        if ($errorCount > 0) {
            $this->warnings[] = sprintf($this->l('%d erreur(s) lors de l\'export'), $errorCount);
        }
    }

    /**
     * Obtenir l'ID de la première commande d'un client
     */
    private function getFirstOrderId($customerId)
    {
        $sql = new DbQuery();
        $sql->select('id_order');
        $sql->from('orders');
        $sql->where('id_customer = ' . (int) $customerId);
        $sql->where('current_state != ' . (int) Configuration::get('PS_OS_CANCELED'));
        $sql->where('current_state != ' . (int) Configuration::get('PS_OS_ERROR'));
        $sql->orderBy('date_add ASC');
        $sql->limit(1);

        return (int) Db::getInstance()->getValue($sql);
    }

    /**
     * Render de la vue
     */
    public function renderView()
    {
        $customerId = (int) Tools::getValue('id_customer');
        $customer = new Customer($customerId);

        if (!Validate::isLoadedObject($customer)) {
            $this->errors[] = $this->l('Client introuvable');
            return;
        }

        $this->context->smarty->assign([
            'customer' => $customer,
            'orderCount' => $this->getCustomerOrderCount($customerId),
            'export_dir' => _PS_MODULE_DIR_ . 'export_clients/exports/',
        ]);

        return parent::renderView();
    }

    /**
     * Définir le titre de la page
     */
    public function initPageHeaderToolbar()
    {
        parent::initPageHeaderToolbar();

        $this->page_header_toolbar_title = $this->l('Export Clients');
    }
}
