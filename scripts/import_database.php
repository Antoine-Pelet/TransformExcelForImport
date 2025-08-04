<?php

$res = 0;
if (!$res && !empty($_SERVER["CONTEXT_DOCUMENT_ROOT"])) $res = @include $_SERVER["CONTEXT_DOCUMENT_ROOT"] . "/main.inc.php";
if (!$res && file_exists("../main.inc.php")) $res = @include "../main.inc.php";
if (!$res && file_exists("../../main.inc.php")) $res = @include "../../main.inc.php";
if (!$res && file_exists("../../../main.inc.php")) $res = @include "../../../main.inc.php";
if (!$res) die("Include of main fails");

require_once DOL_DOCUMENT_ROOT . '/product/class/product.class.php';
require_once DOL_DOCUMENT_ROOT . '/bom/class/bom.class.php';
require_once DOL_DOCUMENT_ROOT . '/fourn/class/fournisseur.product.class.php';
require_once DOL_DOCUMENT_ROOT . '/categories/class/categorie.class.php';



if (empty($user->rights->TransformExcelForImport->transformexcelforimport->write)) {
    accessforbidden();
}

$langs->load("products");

$import_type = GETPOST('import_type', 'alpha');
$rows = $_SESSION['excel_import'] ?? [];
$selectedRows = $_POST['selected_rows'] ?? [];

if (empty($rows)) {
    setEventMessages("Aucune donnée à importer.", null, 'errors');
    header("Location: " . $_SERVER["HTTP_REFERER"]);
    exit;
}

$tva_tx = 20;
$entity = $conf->entity ?? 1;
$pmp = 0;
$liste_add = [];

$db->begin();

try {
    foreach ($rows as $row) {
        $ref            = trim($row['ref'] ?? '');
        $label          = trim($row['label'] ?? '');
        $description    = trim($row['description'] ?? '');
        $qty            = (int)($row['qty'] ?? 1);
        $Fabricant_id = (int)($row['Fabricant_id'] ?? 0);
        $fk_bom         = (int)($row['fk_bom'] ?? 0);
        $product_id     = (int)($row['fk_product'] ?? 0);
        $index          = $row['index'] ?? null;
        $type           = trim($row['type'] ?? '');

        // Vérifie si la ligne est sélectionnée 
        if (!empty($selectedRows) && !in_array($index, $selectedRows)) {
            continue;
        }

        // === Cas 1 : Ajout de produit + tag + fournisseur 
        if ($import_type === 'add_product_supplier') {
            if (empty($ref) || empty($label)) continue;
            if (in_array($ref, $liste_add)) continue;

            // === Création du produit ===
            $product = new Product($db);
            $product->ref            = $ref;
            $product->label          = $label;
            $product->description    = $description;
            $product->tva_tx         = $tva_tx;
            $product->entity         = $entity;
            $product->type           = 0;
            $product->status_buy     = 1;
            $product->status         = 0;
            $product->pmp            = $pmp;
            $product->fk_user_author = $user->id;

            $res = $product->create($user);
            if ($res <= 0) {
                throw new Exception("Erreur création produit [$ref] : " . $product->error);
            }

            $product_id = (int)$product->id;
            if (empty($product_id)) {
                throw new Exception("Produit [$ref] créé mais ID nul !");
            }

            // === Association catégorie ===
            if (!empty($type)) {
                $sql = "SELECT rowid, entity FROM " . MAIN_DB_PREFIX . "categorie WHERE label = '" . $db->escape($type) . "'";
                $resql = $db->query($sql);
                if ($resql) {
                    $obj = $db->fetch_object($resql);
                    if ($obj && $obj->rowid > 0) {
                        $categorie_id = (int)$obj->rowid;

                        $sql = "INSERT INTO " . MAIN_DB_PREFIX . "categorie_product (fk_categorie, fk_product)
                        VALUES (" . ((int)$categorie_id) . ", " . ((int)$product_id) . ")";
                        $resql = $db->query($sql);
                        if (!$resql) {
                            throw new Exception("Erreur insert categorie_product : " . $db->lasterror());
                        }
                    } else {
                        throw new Exception("Catégorie [$type] introuvable.");
                    }
                } else {
                    throw new Exception("Erreur requête catégorie [$type] : " . $db->lasterror());
                }
            }

            // === Association fournisseur ===
            if ($Fabricant_id <= 0) {  
                throw new Exception("Aucun fournisseur sélectionné pour le produit [$ref].");
            }

            $sql = "INSERT INTO llx_product_fournisseur_price (fk_product, fk_soc, ref_fourn, quantity, price, entity, multicurrency_code, multicurrency_tx, multicurrency_unitprice, multicurrency_price, status)
                    VALUES (" . ((int)$product_id) . ", " . ((int)$Fabricant_id) . ", '" . $db->escape($ref) . "', 1, 1, " . ((int)$entity) . ", 'EUR', 1, 1, 1, 1)";
                
            $resql = $db->query($sql);
            if (!$resql) {
                throw new Exception("Erreur insert product_fournisseur_price [$ref] : " . $db->lasterror());
            }

            $liste_add[] = $ref;
        }

        // === Cas 2 : Ajout à une nomenclature (BOM) ===
        elseif ($import_type === 'add_BOM') {
            if ($fk_bom <= 0 || empty($ref) || $qty <= 0) continue;
            $position = (int)($row['position'] ?? 0);
            // Récupération de l'ID produit via sa référence
            $sql = "SELECT rowid FROM " . MAIN_DB_PREFIX . "product WHERE ref = '" . $db->escape($ref) . "'";
            $resql = $db->query($sql);
            $prod = $db->fetch_object($resql);

            if ($prod) {
                $product_id = (int) $prod->rowid;

                // Vérification si le produit existe déjà dans la BOM
                $sql = "SELECT rowid, qty FROM " . MAIN_DB_PREFIX . "bom_bomline 
                        WHERE fk_bom = " . ((int) $fk_bom) . " 
                        AND fk_product = " . $product_id;
                $resql2 = $db->query($sql);
                $existing = $db->fetch_object($resql2);

                if ($existing) {
                    if ($existing->qty === $qty) {
                        continue;
                    }
                    $bomline_rowid = (int) $existing->rowid;

                    // Mise à jour de la quantité existante
                    $sql = "UPDATE " . MAIN_DB_PREFIX . "bom_bomline 
                            SET qty = " . ((int) $qty) . " 
                            WHERE rowid = " . $bomline_rowid;
                    $resql3 = $db->query($sql);

                    if (!$resql3) {
                        throw new Exception("Erreur mise à jour ligne BOM [$ref] : " . $db->lasterror());
                    }
                } else {
                    // Ligne BOM inexistante → création
                    $bomline = new BOMLine($db);
                    $bomline->fk_bom = $fk_bom;
                    $bomline->fk_product = $product_id;
                    $bomline->qty = $qty;
                    $bomline->position = $position;

                    $res = $bomline->create($user);
                    if ($res <= 0) {
                        throw new Exception("Erreur création ligne BOM [$ref] : " . $bomline->error);
                    }
                }
            } else {
                throw new Exception("Produit [$ref] introuvable.");
            }
        }
    }

    $db->commit();
    setEventMessages("Importation terminée avec succès.", null, 'mesgs');
    header("Location: " . dol_buildpath('/custom/transformexcelforimport/transformexcelforimportindex.php', 1));
    exit;
} catch (Exception $e) {
    $db->rollback();
    setEventMessages("Erreur durant l'import : " . $e->getMessage(), null, 'errors');
    header("Location: " . dol_buildpath('/custom/transformexcelforimport/transformexcelforimportindex.php', 1));
    exit;
}
