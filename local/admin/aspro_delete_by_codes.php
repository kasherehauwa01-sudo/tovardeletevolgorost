<?php

use PhpOffice\PhpSpreadsheet\IOFactory;

require_once $_SERVER['DOCUMENT_ROOT'] . '/bitrix/modules/main/include/prolog_admin_before.php';

global $APPLICATION;
global $USER;

$APPLICATION->SetTitle('ASPRO: массовая деактивация/удаление по CODE');

if (!\Bitrix\Main\Loader::includeModule('iblock')) {
    require $_SERVER['DOCUMENT_ROOT'] . '/bitrix/modules/main/include/prolog_admin_after.php';
    echo '<div class="adm-info-message-wrap"><div class="adm-info-message">'
        . htmlspecialcharsbx('Ошибка: не удалось подключить модуль iblock.')
        . '</div></div>';
    require $_SERVER['DOCUMENT_ROOT'] . '/bitrix/modules/main/include/epilog_admin.php';
    return;
}

const ASPRO_DELETE_CODES_MAX = 5000;

/**
 * Экранирует строку для безопасного вывода в HTML.
 */
function asproEsc($value)
{
    return htmlspecialcharsbx((string)$value);
}

/**
 * Возвращает путь к лог-файлу и создает директорию при необходимости.
 */
function asproGetLogPath()
{
    $dir = $_SERVER['DOCUMENT_ROOT'] . '/bitrix/logs';
    if (!is_dir($dir)) {
        mkdir($dir, BX_DIR_PERMISSIONS, true);
    }

    return $dir . '/aspro_delete_by_codes.log';
}

/**
 * Запись технического лога без персональных данных (кроме user_id).
 */
function asproWriteLog($userId, $mode, $iblockId, $codesCount, $result)
{
    $line = sprintf(
        "[%s] user_id=%s mode=%s iblock_id=%s codes=%s result=%s\n",
        date('Y-m-d H:i:s'),
        (int)$userId,
        preg_replace('/[^a-z\-]/', '', (string)$mode),
        (int)$iblockId,
        (int)$codesCount,
        preg_replace('/[\r\n]+/', ' ', (string)$result)
    );

    file_put_contents(asproGetLogPath(), $line, FILE_APPEND);
}

/**
 * Проверка прав пользователя на изменение элементов в инфоблоке.
 */
function asproHasIblockWriteAccess($iblockId)
{
    global $USER;

    if ($USER && $USER->IsAdmin()) {
        return true;
    }

    return CIBlockRights::UserHasRightTo($iblockId, $iblockId, 'iblock_edit');
}

$errors = [];
$rows = [];
$summary = [
    'sourceRows' => 0,
    'uniqueCodes' => 0,
    'found' => 0,
    'notFound' => 0,
    'duplicates' => 0,
    'deactivated' => 0,
    'deleted' => 0,
    'skipped' => 0,
];

$mode = isset($_POST['mode']) ? (string)$_POST['mode'] : 'dry-run';
$iblockId = isset($_POST['iblock_id']) ? (int)$_POST['iblock_id'] : 0;
$confirm = isset($_POST['confirm']) && $_POST['confirm'] === 'Y';
$isPost = $_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['run']);

if ($isPost) {
    if (!check_bitrix_sessid()) {
        $errors[] = 'Сессия истекла. Обновите страницу и попробуйте снова.';
    }

    if (!in_array($mode, ['dry-run', 'deactivate', 'delete'], true)) {
        $errors[] = 'Неизвестный режим выполнения.';
    }

    if ($iblockId <= 0) {
        $errors[] = 'Укажите корректный IBLOCK_ID.';
    } elseif (!asproHasIblockWriteAccess($iblockId)) {
        $errors[] = 'Недостаточно прав. Нужны права на запись в инфоблок или права администратора.';
    }

    if (!isset($_FILES['xlsx_file']) || (int)$_FILES['xlsx_file']['error'] !== UPLOAD_ERR_OK) {
        $errors[] = 'Не удалось загрузить файл XLSX.';
    }

    if (($mode === 'deactivate' || $mode === 'delete') && !$confirm) {
        $errors[] = 'Для выполнения изменений необходимо подтвердить чекбокс "Я понимаю последствия".';
    }

    $uploadedPath = '';
    if (!$errors) {
        $originalName = (string)$_FILES['xlsx_file']['name'];
        $extension = mb_strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
        if ($extension !== 'xlsx') {
            $errors[] = 'Допустим только файл формата .xlsx.';
        }
    }

    if (!$errors) {
        $autoloadPath = $_SERVER['DOCUMENT_ROOT'] . '/local/vendor/autoload.php';
        if (!file_exists($autoloadPath)) {
            $errors[] = 'Библиотека PhpSpreadsheet не найдена. Установите: composer require phpoffice/phpspreadsheet';
        } else {
            require_once $autoloadPath;
            if (!class_exists(IOFactory::class)) {
                $errors[] = 'PhpSpreadsheet установлен некорректно. Проверьте /local/vendor/autoload.php';
            }
        }
    }

    if (!$errors) {
        $tmpDir = $_SERVER['DOCUMENT_ROOT'] . '/upload/tmp';
        if (!is_dir($tmpDir) && !mkdir($tmpDir, BX_DIR_PERMISSIONS, true)) {
            $errors[] = 'Не удалось создать директорию /upload/tmp для временных файлов.';
        } else {
            $uploadedPath = $tmpDir . '/aspro_delete_' . uniqid('', true) . '.xlsx';
            if (!move_uploaded_file($_FILES['xlsx_file']['tmp_name'], $uploadedPath)) {
                $errors[] = 'Не удалось переместить загруженный файл во временную директорию.';
            }
        }
    }

    if (!$errors) {
        try {
            $spreadsheet = IOFactory::load($uploadedPath);
            $sheet = $spreadsheet->getActiveSheet();
            $highestRow = (int)$sheet->getHighestDataRow('A');

            $rawCodes = [];
            for ($rowIndex = 1; $rowIndex <= $highestRow; $rowIndex++) {
                $value = trim((string)$sheet->getCell('A' . $rowIndex)->getCalculatedValue());
                if ($value !== '') {
                    $rawCodes[] = $value;
                }
            }

            $summary['sourceRows'] = count($rawCodes);

            $codes = array_values(array_unique($rawCodes));
            $summary['uniqueCodes'] = count($codes);

            if ($summary['uniqueCodes'] > ASPRO_DELETE_CODES_MAX) {
                $errors[] = 'Превышен лимит уникальных кодов за запуск: ' . ASPRO_DELETE_CODES_MAX . '.';
            }

            if (!$errors) {
                foreach ($codes as $code) {
                    $matches = [];
                    $dbRes = CIBlockElement::GetList(
                        [],
                        [
                            'IBLOCK_ID' => $iblockId,
                            '=CODE' => $code,
                            'CHECK_PERMISSIONS' => 'N',
                        ],
                        false,
                        false,
                        ['ID', 'IBLOCK_ID', 'NAME', 'CODE', 'ACTIVE']
                    );

                    while ($item = $dbRes->Fetch()) {
                        $matches[] = $item;
                    }

                    $matchCount = count($matches);
                    if ($matchCount === 0) {
                        $summary['notFound']++;
                        $rows[] = [
                            'code' => $code,
                            'status' => 'не найдено',
                            'count' => 0,
                            'id' => '-',
                            'name' => '-',
                            'active' => '-',
                            'editLink' => '',
                            'actionResult' => '—',
                        ];
                        continue;
                    }

                    if ($matchCount > 1) {
                        $summary['duplicates']++;
                        foreach ($matches as $item) {
                            $rows[] = [
                                'code' => $code,
                                'status' => 'дубликат',
                                'count' => $matchCount,
                                'id' => $item['ID'],
                                'name' => $item['NAME'],
                                'active' => $item['ACTIVE'],
                                'editLink' => '/bitrix/admin/iblock_element_edit.php?IBLOCK_ID=' . (int)$item['IBLOCK_ID'] . '&type=&ID=' . (int)$item['ID'] . '&lang=' . LANGUAGE_ID,
                                'actionResult' => 'Пропущено: найдено несколько элементов с одинаковым CODE.',
                            ];
                        }
                        continue;
                    }

                    $summary['found']++;
                    $item = $matches[0];
                    $actionResult = 'Ничего не выполнялось (dry-run).';

                    if ($mode === 'deactivate') {
                        $el = new CIBlockElement();
                        if ($item['ACTIVE'] === 'N') {
                            $summary['skipped']++;
                            $actionResult = 'Пропущено: элемент уже неактивен.';
                        } elseif ($el->Update((int)$item['ID'], ['ACTIVE' => 'N'])) {
                            $summary['deactivated']++;
                            $item['ACTIVE'] = 'N';
                            $actionResult = 'Успешно деактивирован.';
                        } else {
                            $summary['skipped']++;
                            $actionResult = 'Ошибка деактивации: ' . (string)$el->LAST_ERROR;
                        }
                    } elseif ($mode === 'delete') {
                        if ($item['ACTIVE'] !== 'N') {
                            $summary['skipped']++;
                            $actionResult = 'Удаление запрещено: элемент активен. Сначала деактивируйте.';
                        } elseif (CIBlockElement::Delete((int)$item['ID'])) {
                            $summary['deleted']++;
                            $actionResult = 'Успешно удален.';
                        } else {
                            $summary['skipped']++;
                            $actionResult = 'Ошибка удаления: операция CIBlockElement::Delete вернула false.';
                        }
                    }

                    $rows[] = [
                        'code' => $code,
                        'status' => 'найдено',
                        'count' => 1,
                        'id' => $item['ID'],
                        'name' => $item['NAME'],
                        'active' => $item['ACTIVE'],
                        'editLink' => '/bitrix/admin/iblock_element_edit.php?IBLOCK_ID=' . (int)$item['IBLOCK_ID'] . '&type=&ID=' . (int)$item['ID'] . '&lang=' . LANGUAGE_ID,
                        'actionResult' => $actionResult,
                    ];
                }
            }
        } catch (\Throwable $e) {
            $errors[] = 'Ошибка чтения XLSX: ' . $e->getMessage();
        }
    }

    if (!empty($uploadedPath) && file_exists($uploadedPath)) {
        unlink($uploadedPath);
    }

    $resultText = $errors ? ('error: ' . implode(' | ', $errors)) : sprintf(
        'ok found=%d not_found=%d duplicates=%d deactivated=%d deleted=%d skipped=%d',
        $summary['found'],
        $summary['notFound'],
        $summary['duplicates'],
        $summary['deactivated'],
        $summary['deleted'],
        $summary['skipped']
    );

    asproWriteLog(
        $USER ? (int)$USER->GetID() : 0,
        $mode,
        $iblockId,
        $summary['uniqueCodes'],
        $resultText
    );
}

require $_SERVER['DOCUMENT_ROOT'] . '/bitrix/modules/main/include/prolog_admin_after.php';
?>
<form method="post" enctype="multipart/form-data">
    <?= bitrix_sessid_post(); ?>
    <table class="adm-detail-content-table edit-table">
        <tr>
            <td width="40%"><label for="iblock_id"><b>IBLOCK_ID каталога (обязательно)</b></label></td>
            <td width="60%"><input type="number" min="1" name="iblock_id" id="iblock_id" value="<?= asproEsc($iblockId > 0 ? $iblockId : ''); ?>" required></td>
        </tr>
        <tr>
            <td><label for="xlsx_file"><b>Файл XLSX</b></label></td>
            <td><input type="file" name="xlsx_file" id="xlsx_file" accept=".xlsx" required></td>
        </tr>
        <tr>
            <td><b>Режим</b></td>
            <td>
                <label><input type="radio" name="mode" value="dry-run" <?= $mode === 'dry-run' ? 'checked' : ''; ?>> Только отчет (dry-run)</label><br>
                <label><input type="radio" name="mode" value="deactivate" <?= $mode === 'deactivate' ? 'checked' : ''; ?>> Снять с публикации (ACTIVE = N)</label><br>
                <label><input type="radio" name="mode" value="delete" <?= $mode === 'delete' ? 'checked' : ''; ?>> Удалить</label>
            </td>
        </tr>
        <tr>
            <td><b>Подтверждение</b></td>
            <td>
                <label>
                    <input type="checkbox" name="confirm" value="Y" <?= $confirm ? 'checked' : ''; ?>>
                    Я понимаю последствия
                </label>
            </td>
        </tr>
    </table>
    <input type="submit" name="run" value="Запустить" class="adm-btn-save">
</form>
<?php if ($errors): ?>
    <div class="adm-info-message-wrap adm-info-message-red">
        <div class="adm-info-message">
            <?php foreach ($errors as $error): ?>
                <div><?= asproEsc($error); ?></div>
            <?php endforeach; ?>
        </div>
    </div>
<?php endif; ?>

<?php if ($isPost && !$errors): ?>
    <h3>Сводка</h3>
    <table class="adm-list-table">
        <tr class="adm-list-table-header">
            <td class="adm-list-table-cell">Показатель</td>
            <td class="adm-list-table-cell">Значение</td>
        </tr>
        <tr><td class="adm-list-table-cell">Строк в файле (непустых)</td><td class="adm-list-table-cell"><?= asproEsc($summary['sourceRows']); ?></td></tr>
        <tr><td class="adm-list-table-cell">Уникальных кодов</td><td class="adm-list-table-cell"><?= asproEsc($summary['uniqueCodes']); ?></td></tr>
        <tr><td class="adm-list-table-cell">Найдено (уникальное совпадение)</td><td class="adm-list-table-cell"><?= asproEsc($summary['found']); ?></td></tr>
        <tr><td class="adm-list-table-cell">Не найдено</td><td class="adm-list-table-cell"><?= asproEsc($summary['notFound']); ?></td></tr>
        <tr><td class="adm-list-table-cell">Коды с дубликатами в инфоблоке</td><td class="adm-list-table-cell"><?= asproEsc($summary['duplicates']); ?></td></tr>
        <tr><td class="adm-list-table-cell">Деактивировано</td><td class="adm-list-table-cell"><?= asproEsc($summary['deactivated']); ?></td></tr>
        <tr><td class="adm-list-table-cell">Удалено</td><td class="adm-list-table-cell"><?= asproEsc($summary['deleted']); ?></td></tr>
        <tr><td class="adm-list-table-cell">Пропущено из-за ошибок/условий</td><td class="adm-list-table-cell"><?= asproEsc($summary['skipped']); ?></td></tr>
    </table>

    <h3>Детальный отчет</h3>
    <table class="adm-list-table" style="width: 100%;">
        <tr class="adm-list-table-header">
            <td class="adm-list-table-cell">Код из файла</td>
            <td class="adm-list-table-cell">Статус</td>
            <td class="adm-list-table-cell">Кол-во совпадений</td>
            <td class="adm-list-table-cell">ID</td>
            <td class="adm-list-table-cell">Название</td>
            <td class="adm-list-table-cell">ACTIVE</td>
            <td class="adm-list-table-cell">Ссылка на редактирование</td>
            <td class="adm-list-table-cell">Результат действия</td>
        </tr>
        <?php foreach ($rows as $row): ?>
            <tr>
                <td class="adm-list-table-cell"><?= asproEsc($row['code']); ?></td>
                <td class="adm-list-table-cell"><?= asproEsc($row['status']); ?></td>
                <td class="adm-list-table-cell"><?= asproEsc($row['count']); ?></td>
                <td class="adm-list-table-cell"><?= asproEsc($row['id']); ?></td>
                <td class="adm-list-table-cell"><?= asproEsc($row['name']); ?></td>
                <td class="adm-list-table-cell"><?= asproEsc($row['active']); ?></td>
                <td class="adm-list-table-cell">
                    <?php if (!empty($row['editLink'])): ?>
                        <a href="<?= asproEsc($row['editLink']); ?>" target="_blank"><?= asproEsc('Открыть'); ?></a>
                    <?php else: ?>
                        <?= asproEsc('-'); ?>
                    <?php endif; ?>
                </td>
                <td class="adm-list-table-cell"><?= asproEsc($row['actionResult']); ?></td>
            </tr>
        <?php endforeach; ?>
    </table>
<?php endif; ?>

<?php require $_SERVER['DOCUMENT_ROOT'] . '/bitrix/modules/main/include/epilog_admin.php'; ?>
