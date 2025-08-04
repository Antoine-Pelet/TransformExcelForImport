<?php

$res = 0;
if (!$res && !empty($_SERVER["CONTEXT_DOCUMENT_ROOT"])) $res = @include $_SERVER["CONTEXT_DOCUMENT_ROOT"] . "/main.inc.php";
if (!$res && file_exists("../main.inc.php")) $res = @include "../main.inc.php";
if (!$res && file_exists("../../main.inc.php")) $res = @include "../../main.inc.php";
if (!$res && file_exists("../../../main.inc.php")) $res = @include "../../../main.inc.php";
if (!$res) die("Include of main fails");
if (empty($user->id)) accessforbidden();


// Vérification que l'utilisateur est bien chargé
if (!is_object($user) || empty($user->id)) {
    accessforbidden('Utilisateur non connecté ou invalide.');
}

// Vérification des droits
if (empty($user->rights->TransformExcelForImport->transformexcelforimport->write)) {
    accessforbidden('Droits insuffisants pour accéder à ce module.');
}

require_once DOL_DOCUMENT_ROOT . '/core/class/html.form.class.php';
require_once DOL_DOCUMENT_ROOT . '/fourn/class/fournisseur.class.php';
require_once DOL_DOCUMENT_ROOT . '/bom/class/bom.class.php';
require_once DOL_DOCUMENT_ROOT . '/core/lib/admin.lib.php';

$langs->loadLangs(array("transformexcelforimport@transformexcelforimport"));



$form = new Form($db);

// Récupération des fournisseurs
$suppliers = [];
$sql = "SELECT rowid, nom FROM " . MAIN_DB_PREFIX . "societe WHERE fournisseur = 1";
$resql = $db->query($sql);
if ($resql) while ($obj = $db->fetch_object($resql)) $suppliers[] = $obj;

// Récupération des BOM
$boms = [];
$sql = "SELECT rowid, label FROM " . MAIN_DB_PREFIX . "bom_bom where status = 0";
$resql = $db->query($sql);
if ($resql) while ($obj = $db->fetch_object($resql)) $boms[] = $obj;

$csrf_token = newToken('preview_excel');

llxHeader("", $langs->trans("transformexcelforimportArea"));
print load_fiche_titre($langs->trans("Transforme excel pour import"), '', 'transformexcelforimport.png@transformexcelforimport');
?>

<form id="uploadForm" enctype="multipart/form-data" method="POST" action="scripts/product_fourn.php">
    <input type="hidden" name="token" value="<?php echo $csrf_token; ?>">

    <table class="noborder" width="100%">
        <tr>
            <td width="20%">Fichier Excel :</td>
            <td><input type="file" name="excel_file" id="excel_file" required></td>
        </tr>
        <tr>
            <td>Type d'import :</td>
            <td>
                <select name="script" id="scriptType" required>
                    <option value="">-- Choisir --</option>
                    <option value="add_product">Ajouter des produits</option>
                    <option value="add_BOM">Ajouter à une BOM</option>
                </select>
            </td>
        </tr>
        <tr class="supplier-row" style="display:none;">
            <td>Fournisseur :</td>
            <td>
                <select name="supplier_id" id="supplier_id">
                    <?php foreach ($suppliers as $s) echo "<option value='{$s->rowid}'>{$s->nom}</option>"; ?>
                </select>
            </td>
        </tr>
        <tr class="bom-row" style="display:none;">
            <td>BOM :</td>
            <td>
                <select name="bom_id" id="bom_id">
                    <?php foreach ($boms as $b) echo "<option value='{$b->rowid}'>{$b->label}</option>"; ?>
                </select>
            </td>
        </tr>
    </table>
    <div class="center">
        <button type="submit" class="button">Analyser le fichier</button>
    </div>
</form>

<script>
    document.getElementById('scriptType').addEventListener('change', function() {
        const type = this.value;
        document.querySelector('.supplier-row').style.display = (type === 'add_product') ? '' : 'none';
        document.querySelector('.bom-row').style.display = (type === 'add_BOM') ? '' : 'none';
    });
</script>
<?php if (!empty($mesg)) : ?>
    <script>
        document.addEventListener('DOMContentLoaded', function() {
            Swal.fire({
                title: "<?php echo $mesgtype === 'errors' ? 'Erreur' : 'Succès'; ?>",
                text: "<?php echo is_array($mesg) ? implode('\n', $mesg) : $mesg; ?>",
                icon: "<?php echo $mesgtype === 'errors' ? 'error' : 'success'; ?>",
                confirmButtonText: 'OK'
            });
        });
    </script>
<?php endif; ?>


<?php
llxFooter();
$db->close();
