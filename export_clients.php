<?php
/**
 * Module Export Clients
 *
 * @author Consogarage
 * @copyright 2025 Consogarage
 * @license Proprietary
 */

if (!defined('_PS_VERSION_')) {
    exit;
}

require_once dirname(__FILE__) . '/classes/ExportClientsGenerator.php';

class Export_Clients extends Module
{
    public function __construct()
    {
        $this->name = 'export_clients';
        $this->tab = 'administration';
        $this->version = '1.0.0';
        $this->author = 'Consogarage';
        $this->need_instance = 1;
        $this->ps_versions_compliancy = [
            'min' => '8.0.0',
            'max' => '8.99.99',
        ];
        $this->bootstrap = true;

        parent::__construct();

        $this->displayName = $this->l('Export Clients');
        $this->description = $this->l('Export automatique des informations clients lors de leur première commande au format XLSM');

        $this->confirmUninstall = $this->l('Êtes-vous sûr de vouloir désinstaller ce module ?');
    }

    /**
     * Installation du module
     */
    public function install()
    {
        if (
            !parent::install() ||
            !$this->registerHook('actionValidateOrder') ||
            !$this->installTab()
        ) {
            return false;
        }

        // Créer le dossier d'export si nécessaire
        $exportDir = _PS_MODULE_DIR_ . $this->name . '/exports';
        if (!file_exists($exportDir)) {
            mkdir($exportDir, 0755, true);
        }

        // Créer le fichier .htaccess pour protéger les exports
        $htaccess = $exportDir . '/.htaccess';
        if (!file_exists($htaccess)) {
            file_put_contents($htaccess, "deny from all\n");
        }

        return true;
    }

    /**
     * Désinstallation du module
     */
    public function uninstall()
    {
        if (
            !parent::uninstall() ||
            !$this->uninstallTab()
        ) {
            return false;
        }

        return true;
    }

    /**
     * Installation de l'onglet d'administration
     */
    private function installTab()
    {
        $tab = new Tab();
        $tab->active = 1;
        $tab->class_name = 'AdminExportClients';
        $tab->name = [];
        foreach (Language::getLanguages(true) as $lang) {
            $tab->name[$lang['id_lang']] = 'Export Clients';
        }
        $tab->id_parent = (int) Tab::getIdFromClassName('AdminParentOrders');
        $tab->module = $this->name;

        return $tab->add();
    }

    /**
     * Désinstallation de l'onglet d'administration
     */
    private function uninstallTab()
    {
        $idTab = (int) Tab::getIdFromClassName('AdminExportClients');
        if ($idTab) {
            $tab = new Tab($idTab);
            return $tab->delete();
        }

        return true;
    }

    /**
     * Hook exécuté lors de la validation d'une commande
     *
     * @param array $params
     */
    public function hookActionValidateOrder($params)
    {
        if (!isset($params['order']) || !isset($params['customer'])) {
            return;
        }

        $order = $params['order'];
        $customer = $params['customer'];

        // Vérifier si c'est la première commande du client
        if ($this->isFirstOrder($customer->id)) {
            try {
                $generator = new ExportClientsGenerator();
                $generator->generateExport($customer->id, $order->id);
            } catch (Exception $e) {
                PrestaShopLogger::addLog(
                    'Export Clients: Erreur lors de la génération de l\'export - ' . $e->getMessage(),
                    3,
                    null,
                    'Customer',
                    $customer->id
                );
            }
        }
    }

    /**
     * Vérifier si c'est la première commande du client
     *
     * @param int $customerId
     * @return bool
     */
    private function isFirstOrder($customerId)
    {
        $sql = new DbQuery();
        $sql->select('COUNT(*)');
        $sql->from('orders');
        $sql->where('id_customer = ' . (int) $customerId);
        $sql->where('current_state != ' . (int) Configuration::get('PS_OS_CANCELED'));
        $sql->where('current_state != ' . (int) Configuration::get('PS_OS_ERROR'));

        return (int) Db::getInstance()->getValue($sql) === 1;
    }

    /**
     * Page de configuration du module
     */
    public function getContent()
    {
        $output = '';

        if (Tools::isSubmit('submit' . $this->name)) {
            $output .= $this->displayConfirmation($this->l('Configuration mise à jour'));
        }

        return $output . $this->displayForm();
    }

    /**
     * Formulaire de configuration
     */
    public function displayForm()
    {
        $fieldsForm = [
            'form' => [
                'legend' => [
                    'title' => $this->l('Configuration'),
                    'icon' => 'icon-cogs',
                ],
                'input' => [
                    [
                        'type' => 'switch',
                        'label' => $this->l('Activer l\'export automatique'),
                        'name' => 'EXPORT_CLIENTS_ENABLED',
                        'is_bool' => true,
                        'desc' => $this->l('Activer ou désactiver l\'export automatique lors de la première commande'),
                        'values' => [
                            [
                                'id' => 'active_on',
                                'value' => 1,
                                'label' => $this->l('Activé'),
                            ],
                            [
                                'id' => 'active_off',
                                'value' => 0,
                                'label' => $this->l('Désactivé'),
                            ],
                        ],
                    ],
                ],
                'submit' => [
                    'title' => $this->l('Enregistrer'),
                    'class' => 'btn btn-default pull-right',
                ],
            ],
        ];

        $helper = new HelperForm();
        $helper->show_toolbar = false;
        $helper->table = $this->table;
        $helper->module = $this;
        $helper->default_form_language = (int) Configuration::get('PS_LANG_DEFAULT');
        $helper->allow_employee_form_lang = Configuration::get('PS_BO_ALLOW_EMPLOYEE_FORM_LANG', 0);
        $helper->identifier = $this->identifier;
        $helper->submit_action = 'submit' . $this->name;
        $helper->currentIndex = $this->context->link->getAdminLink('AdminModules', false)
            . '&configure=' . $this->name . '&tab_module=' . $this->tab . '&module_name=' . $this->name;
        $helper->token = Tools::getAdminTokenLite('AdminModules');
        $helper->tpl_vars = [
            'fields_value' => [
                'EXPORT_CLIENTS_ENABLED' => Configuration::get('EXPORT_CLIENTS_ENABLED', true),
            ],
            'languages' => $this->context->controller->getLanguages(),
            'id_language' => $this->context->language->id,
        ];

        return $helper->generateForm([$fieldsForm]);
    }
}
