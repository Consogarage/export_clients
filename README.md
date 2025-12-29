# export_clients
Module Prestashop (8.2.2) d'export clients

## Principe de fonctionnement
Lorsqu'un client effectue sa premère commande sur le site, le module génère un fichier XLSM contenant les informations suivantes :
- Coordonnées
  - Titre
  - Raison sociale
  - Adresse
  - Code postal
  - Ville
  - Pays
  - Téléphone
  - Bureau
  - Télécopie
  - Portable
  - Adresse électronique (email)
  - Site web
- Détail
  - Classification
    - Code groupe
    - Code famille
    - Collectif
  - Renseignements complémentaires
    - Numéro de SIRET
    - Code NAF
    - N° TVA intracommunautaire
    - N° EDI
    - Nom interlocuteur
    - Commentaires
    - Code TVA
- Paiement
  - Conditions de paiement
    - Mode de paiement
    - Délai de paiement
    - Décalage
    - Type
    - Banque d'affectation
  - Coordonnées bancaires
    - Domiciliation
    - IBAN
    - RIB
    - Code BIC 