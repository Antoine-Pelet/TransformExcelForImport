<?php

$res = 0;
if (!$res && file_exists("../../main.inc.php")) $res = include "../../main.inc.php";
if (!$res && file_exists("../../../main.inc.php")) $res = include "../../../main.inc.php";
if (!$res) die("Include of main fails");

require_once DOL_DOCUMENT_ROOT . '/custom/transformexcelforimport/lib/vendor/autoload.php';
require_once DOL_DOCUMENT_ROOT . '/core/lib/admin.lib.php';
require_once DOL_DOCUMENT_ROOT . '/user/class/user.class.php';

if (empty($user->rights->TransformExcelForImport->transformexcelforimport->write)) {
    accessforbidden();
}

use PhpOffice\PhpSpreadsheet\IOFactory;

global $langs, $db, $conf;

$label = ($_POST['script'] == 'add_product') ? $langs->trans("Ajout produit") : $langs->trans("Ajout à une BOM");
$title = $langs->trans("RésultatDeLImportExcel") . ' : ' . $label;
llxHeader('', $title);


$type_import = $_POST['script'] ?? '';
$fk_bom = $_POST['bom_id'] ?? '';


$type_import = in_array($_POST['script'], ['add_product', 'add_BOM']) ? $_POST['script'] : '';
if ($_POST['script'] == 'add_BOM') $fk_bom = isset($_POST['bom_id']) && is_numeric($_POST['bom_id']) ? (int) $_POST['bom_id'] : 0;


$tmpFile = $_FILES['excel_file']['tmp_name'] ?? null;

print '<div class="fiche">';
print '<div class="titre">' . $langs->trans("Résultat de l'import Excel" . ' : ' . $label) . '</div><br>';


if (!isset($_FILES['excel_file']) || $_FILES['excel_file']['error'] !== UPLOAD_ERR_OK) {
    throw new Exception("Erreur lors de l'upload du fichier Excel.");
}

$allowedMimeTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel'
];

if (!in_array($_FILES['excel_file']['type'], $allowedMimeTypes)) {
    throw new Exception("Le fichier n'est pas un fichier Excel valide.");
}

function contient_url($texte)
{
    return preg_match('/https?:\/\/|www\./i', $texte);
}

function construire_ref($ref, $description, $indice, $Fabricant)
{
    $ref = trim((string) $ref);
    if (contient_url($ref) == '') return null;
    if (strpos($ref, 'ASP') === 0) return $ref . '_' . $indice; //normaliser_ref($ref . '_' . $indice);

    if ($Fabricant === '' || $Fabricant === 'null' || $Fabricant === 'undefined') {
        return $ref; //normaliser_ref($ref);
    } else {
        return $Fabricant . '_' . $ref; //normaliser_ref($Fabricant . '_' . $ref);
    }
}

try {
    $spreadsheet = IOFactory::load($tmpFile);
    $sheet = $spreadsheet->getActiveSheet();
    $rows_validated = [];
    $line_exists = [];

    $starting_line_number = 1;
    $line_counter = 0;

    if ($_POST['script'] == 'add_BOM') {
        if ($_POST['bom_id'] > 0) {
            if (!$db) {
                throw new Exception("Database connection not available.");
            }


            $sql = "SELECT MAX(position) AS maxline 
            FROM " . MAIN_DB_PREFIX . "bom_bomline 
            WHERE fk_bom = " . ((int) $_POST['bom_id']);

            $resql = $db->query($sql);
            if ($resql && ($obj = $db->fetch_object($resql))) {
                $starting_line_number = ((int) $obj->maxline) + 1;
            }
            if ($starting_line_number > 1) {
                $sql = "
                    SELECT p.ref, bbl.qty 
                    FROM " . MAIN_DB_PREFIX . "bom_bomline bbl
                    JOIN " . MAIN_DB_PREFIX . "product p ON p.rowid = bbl.fk_product
                    WHERE bbl.fk_bom = " . ((int) $_POST['bom_id']);

                $resql = $db->query($sql);
                if ($resql) {
                    while ($obj = $db->fetch_object($resql)) {
                        $line_exists[$obj->ref] = $obj->qty;
                    }
                } else {
                    throw new Exception("Erreur lors de la récupération des lignes de la BOM : " . $db->lasterror());
                }
            }
        }
    } elseif ($_POST['script'] == 'add_product') {
        require_once DOL_DOCUMENT_ROOT . '/categories/class/categorie.class.php';
        if (!$db) {
            throw new Exception("Database connection not available.");
        }

        $db->begin();
        $tag_name=['FABRICATION', 'SOUS-TRAITANCE', 'FAB-INTERNE', 'COMMERCE', 'QUINCAILLERIE'];
        for ($i = 0; $i < 5; $i++) {
            $sql = "SELECT * FROM " . MAIN_DB_PREFIX . "llx_categorie where label = '" . $tag_name[$i] . "'";
            $resql = $db->query($sql);
            if (!$resql) {

                $tag = new Categorie($db);
                $tag->label = $tag_name[$i];
                $tag->entity = $conf->entity;
                $tag->fk_user_author = $user->id;

                $res = $tag->create($user);
            }
        }

        $db->commit();
    }


    $headers = [];

    $rows_validated = [];
    foreach ($sheet->getRowIterator() as $i => $row) {
        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false);
        $rowData = [];

        foreach ($cellIterator as $cell) {
            $rowData[] = trim((string)$cell->getValue());
        }

        if ($i === 1) {
            $headers = array_map('strtolower', array_map('trim', $rowData));
            continue;
        }

        $assoc = [];
        foreach ($rowData as $k => $v) {
            if (isset($headers[$k])) {
                $assoc[$headers[$k]] = $v;
            }
        }


        $ref      = $assoc['référence'] ?? $assoc['reference'] ?? '';
        $desc     = $assoc['description'] ?? $assoc['déscription'] ?? '';
        if ($desc === '') {
            $desc = $ref;
        }
        $indice   = $assoc['indice'] ?? '';
        $Fabricant = $assoc['fabricant'] ?? '';
        $type = $assoc['type'] ?? '';


        $Ref = construire_ref($ref, $desc, $indice, $Fabricant);
        if (!$Ref) continue;

        $original = $assoc;
        $qty1 = '';
        foreach ($original as $key => $value) {
            if (preg_match('/quantité/i', $key)) {
                $qty1 = $value;
                break;
            }
        }

        $row = [
            'ref'         => $Ref,
            'label'       => $desc,
            'description' => $desc,
            'qty'         => $qty1,
            'fk_product'  => $assoc['fk_product'] ?? 0,
            'original'    => $assoc,
            'index'       => $i,
            'type'        => $type,
        ];

        if ($_POST['script'] === 'add_product') {
            $row['Fabricant_id'] = $_POST['supplier_id'] ?? 0;
        }
        if ($_POST['script'] === 'add_BOM') {
            $row['fk_bom'] = $_POST['bom_id'] ?? 0;
            $row['position'] = $starting_line_number;
        }

        $rows_validated[] = $row;
    }


    // Vérifie en base si chaque produit existe déjà
    $existing_refs = [];
    if (!empty($rows_validated)) {
        $refs_to_check = array_map(function ($r) {
            global $db;
            if (!$db) {
                throw new Exception("Database connection not available.");
            }
            return "'" . $db->escape($r['ref']) . "'";
        }, $rows_validated);
        if (!$db) {
            throw new Exception("Database connection not available.");
        }
        $sql = "SELECT ref FROM " . MAIN_DB_PREFIX . "product WHERE ref IN (" . implode(',', $refs_to_check) . ")";
        $resql = $db->query($sql);
        if ($resql) {
            while ($obj = $db->fetch_object($resql)) {
                $existing_refs[] = $obj->ref;
            }
        }
    }



    $affichables = 0;
    foreach ($rows_validated as $row) {
        $Ref = $row['ref'];
        $ref_clean = trim($row['ref'] ?? '');
        $qty1 = $row['qty'];
        $desc = $row['description'] ?? '';
        $is_existing = in_array($Ref, $existing_refs);
        $qty_in_bom = $line_exists[$ref_clean] ?? null;
        $is_in_bom = array_key_exists($ref_clean, $line_exists);
        $qty_diff = ($is_in_bom && $qty1 != $qty_in_bom);

        $do_create = false;
        if ($type_import === 'add_product') {
            $do_create = !$is_existing;
        } elseif ($type_import === 'add_BOM') {
            $do_create = true;
        }

        if ($do_create) $affichables = $affichables + 1;
    }

    $ligne_affichee = 0;
    $has_red_line = false;
    $has_orange_line = false;
    $add_product = false;
    $modif_bom = false;


    if (empty($rows_validated)) {
        print '<div class="warning">Aucune ligne valide trouvée après traitement.</div>';
    } else {
        if ($affichables > 0) {
            print '<form method="POST" action="import_database.php" id="importForm">';
            print '<input type="hidden" name="import_type" value="' . dol_escape_htmltag($type_import) . '">';
            print '<table class="noborder liste" width="100%">';

            print '<tr class="liste_titre">';
            if ($type_import === 'add_product') {
                print '<th><input type="checkbox" onclick="toggleAll(this)"></th>';
                print '<th>reference</th>';
                print '<th>Description </th>';
            } else {
                if ($starting_line_number > 1) {
                    print '<th>Quantité pour 1 actuel</th>';
                }
                print '<th>Quantité pour 1</th>';
                print '<th>reference</th>';
                print '<th>Description </th>';
                print '<th>numéro de ligne</th>';
                if ($has_orange_line) {
                    print '<th><input type="checkbox" onclick="toggleAll(this)"></th>';
                }
            }
            print '</tr>';
        }

        foreach ($rows_validated as $row) {
            $original = $row['original'];
            $qty1 = $row['qty'];
            $ref_clean = trim($row['ref'] ?? '');
            $desc = $row['description'] ?? '';
            $Ref = $row['ref'];

            $is_existing = in_array($Ref, $existing_refs);
            $qty_in_bom = $line_exists[$ref_clean] ?? null;
            $is_in_bom = array_key_exists($ref_clean, $line_exists);
            $qty_diff = ($is_in_bom && $qty1 != $qty_in_bom);
            $style = '';
            $show_checkbox = false;
            $do_create = false;
            if ($type_import === 'add_product') {
                if ($is_existing) {
                    $do_create = false;
                } else {
                    $show_checkbox = true;
                    $add_product = true;
                    $do_create = true;
                }
            } elseif ($type_import === 'add_BOM') {
                if (!$is_existing) {
                    $style = ' style="color: red;"';
                    $do_create = true;
                    $has_red_line = true;
                } elseif ($is_in_bom && !$qty_diff) {
                    $style = ' style="opacity: 0.5;"';
                    $do_create = true;
                } elseif ($is_in_bom && $qty_diff) {
                    $style = ' style="color: orange;"';
                    $show_checkbox = true;
                    $modif_bom = true;
                    $has_orange_line = true;
                    $do_create = true;
                } else {
                    $style = ''; // produit en base mais pas dans la BOM
                    $show_checkbox = true;
                    $modif_bom = true;
                    $do_create = true;
                }
            }

            // Saut des lignes non créées
            if (!$do_create) {
                continue;
            }
            if ($is_in_bom) {
                if (!$db) {
                    throw new Exception("Database connection not available.");
                }

                // Étape 1 : récupérer l'ID du produit via sa référence
                $sql = "SELECT rowid FROM " . MAIN_DB_PREFIX . "product WHERE ref = '" . $Ref . "'";


                $resql_product = $db->query($sql);



                if ($resql_product && ($obj_product = $db->fetch_object($resql_product))) {
                    $product_id = (int) $obj_product->rowid;

                    // Étape 2 : récupérer la position depuis bom_bomline
                    $sql = "SELECT position 
                            FROM " . MAIN_DB_PREFIX . "bom_bomline 
                            WHERE fk_bom = " . ((int) $_POST['bom_id']) . " 
                            AND fk_product = " . $product_id;

                    $resql = $db->query($sql);

                    if ($resql && ($obj = $db->fetch_object($resql))) {
                        $current_line_number = (int) $obj->position;
                    } else {
                        throw new Exception("Impossible de récupérer la position de la ligne BOM pour le produit ID $product_id (ref=$ref)");
                    }
                } else {
                    continue;
                }
            } else {
                $current_line_number = $starting_line_number + $ligne_affichee;
                $ligne_affichee++;
            }




            print '<tr' . $style . '>';


            if ($type_import === 'add_product') {

                if ($show_checkbox) {
                    print '<td class="center"><input type="checkbox" name="selected_rows[]" value="' . $row['index'] . '" checked></td>';
                } else {
                    print '<td class="center">—</td>';
                }

                print '<td>' . dol_escape_htmltag($Ref) . '</td>';
                print '<td>' . dol_escape_htmltag($desc) . '</td>';
            } elseif ($type_import === 'add_BOM') {


                if ($starting_line_number > 1) {
                    print '<td>' . dol_escape_htmltag((int)$qty_in_bom ?? '-') . '</td>';
                }

                print '<td>' . dol_escape_htmltag($qty1) . '</td>';
                print '<td>' . dol_escape_htmltag($Ref) . '</td>';
                print '<td>' . dol_escape_htmltag($desc) . '</td>';
                print '<td class="center">' . dol_escape_htmltag($current_line_number) . '</td>';

                if (!$is_in_bom && $is_existing) {
                    // Produit en base mais pas dans la BOM : on l'ajoute automatiquement (non décochable)
                    print '<td class="center">—</td>';
                    print '<input type="hidden" name="selected_rows[]" value="' . $row['index'] . '">';
                } elseif ($show_checkbox) {
                    // Ligne orange : permet à l'utilisateur de choisir
                    print '<td class="center"><input type="checkbox" name="selected_rows[]" value="' . $row['index'] . '" checked></td>';
                } else {
                    print '<td class="center">—</td>';
                }
            }

            print '</tr>';
        }






        print '</table>';
        print '<br><div class="center">';

        $_SESSION['excel_import'] = $rows_validated;


        print '<br><div class="center">';
        print '<input type="hidden" name="token" value="' . newToken() . '">';
        print '<input type="hidden" name="import_id" value="session">';

        // Détermination du type d'import
        if ($_POST['script'] === 'add_BOM') {
            print '<input type="hidden" name="import_type" value="add_BOM">';
        } elseif ($_POST['script'] === 'add_product') {
            print '<input type="hidden" name="import_type" value="add_product_supplier">';
        }

        print '</table>';



        if ($type_import === 'add_BOM' && $has_red_line) {
            print '<div class="error">Import impossible : certaines lignes font référence à des produits inexistants (lignes rouges).</div>';
            print '<div class="warning">Veuillez faire l\'import produit avant de réessayer.</div>';
        } elseif ($type_import === 'add_BOM' && !$modif_bom) {
            print '<div class="warning">Aucunes lignes sont à modifiées/ajoutées par rapport à la BOM existante.</div>';
        } elseif ($type_import === 'add_product' && !$add_product) {
            print '<div class="warning">Aucun produit à ajouter</div>';
        } else {
            print '<div class="center"><input type="submit" class="button" value="' . $langs->trans("Importer") . '"></div>';
        }

        print '</div>';
        print '</form>';
    }
} catch (Exception $e) {
    print '<div class="error">Erreur lecture Excel : ' . dol_escape_htmltag($e->getMessage()) . '</div>';
}

print '<br><div class="center">';
print '<a class="butAction" href="' . DOL_URL_ROOT . '/custom/transformexcelforimport/transformexcelforimportindex.php">← Retour au menu d\'import</a>';
print '</div>';
print '</div>';

llxFooter();
?>

<style>
    table.liste {
        border-collapse: collapse;
    }

    table.liste th,
    table.liste td {
        border: 1px solid #ccc;
        padding: 8px;
        text-align: left;
        vertical-align: top;
    }

    table.liste th {
        background-color: #f0f0f0;
    }

    tr:nth-child(even) {
        background-color: #f9f9f9;
    }
</style>
<script>
    function toggleAll(masterCheckbox) {
        const checkboxes = document.querySelectorAll('input[name="selected_rows[]"]:not([disabled])');
        checkboxes.forEach(checkbox => {
            checkbox.checked = masterCheckbox.checked;
        });
    }
</script>