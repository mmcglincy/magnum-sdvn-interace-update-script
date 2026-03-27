<?php

declare(strict_types=1);

/**
 * Usage:
 *   php move_round_mapper.php <lookup_csv> <target_csv>
 *
 * lookup_csv:
 *   Column 3 = line name
 *   Column 4 = device name
 *   Column 9 = move round
 *
 * target_csv:
 *   Column 1 = interface_name
 *   Column 4 = SRC/DST
 *   Column 6 = device name (replaced)
 *   Column 7 = line name (match key)
 */

if ($argc < 3) {
    fwrite(STDERR, "Usage: php move_round_mapper.php <lookup_csv> <target_csv>\n");
    exit(1);
}

$lookupPath = $argv[1];
$targetPath = $argv[2];

if (!is_file($lookupPath)) {
    fwrite(STDERR, "Lookup CSV not found: {$lookupPath}\n");
    exit(1);
}

if (!is_file($targetPath)) {
    fwrite(STDERR, "Target CSV not found: {$targetPath}\n");
    exit(1);
}

/**
 * Read CSV into array-of-rows.
 *
 * @return array<int, array<int, string|null>>
 */
function readCsvRows(string $path): array
{
    $rows = [];
    $handle = fopen($path, 'rb');
    if ($handle === false) {
        throw new RuntimeException("Unable to open CSV: {$path}");
    }

    while (($row = fgetcsv($handle, 0, ',', '"', '\\')) !== false) {
        $rows[] = $row;
    }

    fclose($handle);
    return $rows;
}

/**
 * Build round map keyed by stable internal round keys.
 *
 * - round labels can be decimals (example: 2.5)
 * - decimal commas are accepted (example: 2,5)
 * - rows without a specified numeric move round are skipped
 * - all specified move rounds are preserved even when mapping fields are empty
 *
 * @param array<int, array<int, string|null>> $lookupRows
 * @return array<string, array{roundLabel: string, lineToDevice: array<string, string>}>
 */
function buildRoundMap(array $lookupRows): array
{
    $roundMap = [];

    foreach ($lookupRows as $row) {
        // Column 9 (index 8): move round
        if (!array_key_exists(8, $row)) {
            continue;
        }

        $roundRaw = trim((string) $row[8]);
        $roundLabel = normalizeRoundLabel($roundRaw);
        if ($roundLabel === '') {
            continue;
        }

        $roundKey = makeRoundKey($roundLabel);
        if (!isset($roundMap[$roundKey])) {
            // Keep every explicitly specified round from column 9.
            $roundMap[$roundKey] = [
                'roundLabel' => $roundLabel,
                'lineToDevice' => [],
            ];
        }

        // Column 3 (index 2): line name
        // Column 4 (index 3): device name
        if (!array_key_exists(2, $row) || !array_key_exists(3, $row)) {
            continue;
        }

        $lineName = trim((string) $row[2]);
        $deviceName = trim((string) $row[3]);
        if ($lineName === '' || $deviceName === '') {
            continue;
        }

        // If duplicate line names exist in same round, later row wins.
        $roundMap[$roundKey]['lineToDevice'][$lineName] = $deviceName;
    }

    return $roundMap;
}

/**
 * Normalize numeric round labels for filenames/sorting.
 *
 * Note: preserves original decimal precision (e.g. 2.0 stays 2.0).
 */
function normalizeRoundLabel(string $roundRaw): string
{
    $value = trim($roundRaw);
    if ($value === '') {
        return '';
    }

    // Handle UTF-8 BOM that can appear in CSV exports.
    $value = preg_replace('/^\xEF\xBB\xBF/u', '', $value) ?? $value;

    // Accept spreadsheet-style decimal commas.
    $value = str_replace(',', '.', $value);
    if (!is_numeric($value)) {
        return '';
    }

    // Remove a leading plus sign while preserving the original precision.
    if (str_starts_with($value, '+')) {
        $value = substr($value, 1);
    }

    return $value === '-0' ? '0' : $value;
}

/**
 * Build an internal key from a normalized round label.
 */
function makeRoundKey(string $roundLabel): string
{
    return 'round::' . $roundLabel;
}

/**
 * Extract normalized round label from internal key.
 */
function roundLabelFromKey(string $roundKey): string
{
    return str_starts_with($roundKey, 'round::') ? substr($roundKey, 7) : $roundKey;
}

/**
 * Get round keys sorted by numeric round label (ascending).
 *
 * @param array<string, array{roundLabel: string, lineToDevice: array<string, string>}> $roundMap
 * @return array<int, string>
 */
function getSortedRoundKeys(array $roundMap): array
{
    $roundKeys = array_keys($roundMap);
    usort(
        $roundKeys,
        static function (string $a, string $b): int {
            $labelA = roundLabelFromKey($a);
            $labelB = roundLabelFromKey($b);
            $valueCompare = (float) $labelA <=> (float) $labelB;
            if ($valueCompare !== 0) {
                return $valueCompare;
            }

            // Deterministic tie-breaker if numeric values are equal.
            return strcmp($labelA, $labelB);
        }
    );

    return $roundKeys;
}

/**
 * Build unique interface_name list from target CSV column 1.
 *
 * @param array<int, array<int, string|null>> $targetRows
 * @return array<int, string>
 */
function getUniqueInterfaceNames(array $targetRows): array
{
    $seen = [];
    $interfaceNames = [];

    foreach ($targetRows as $rowIndex => $row) {
        // Do not use the first row when collecting interface_name values.
        if ($rowIndex === 0) {
            continue;
        }

        if (!array_key_exists(0, $row)) {
            continue;
        }

        $interfaceName = trim((string) $row[0]);
        if (
            $interfaceName === '' ||
            isset($seen[$interfaceName])
        ) {
            continue;
        }

        $seen[$interfaceName] = true;
        $interfaceNames[] = $interfaceName;
    }

    return $interfaceNames;
}

/**
 * Convert interface names to safe filename parts.
 */
function sanitizeFilenamePart(string $value): string
{
    $sanitized = preg_replace('/[^A-Za-z0-9._-]+/', '_', trim($value));
    if ($sanitized === null || $sanitized === '') {
        return 'unknown';
    }

    return $sanitized;
}

function debugEcho(string $message): void
{
    echo "[DEBUG] {$message}\n";
}

/**
 * Apply replacements from a prepared line->device map.
 *
 * Rules:
 * - Only for rows participating in the current cumulative move-round set
 *   (line name exists in effectiveMap), if TOPS Name (column 7) partially
 *   matches ITXR, remove the row.
 * - Only replace when a target row line-name matches, column 4 is SRC, and existing device is ITXE.
 *
 * @param array<int, array<int, string|null>> $targetRows
 * @param array<string, string> $effectiveMap
 * @return array<int, array<int, string|null>>
 */
function rowHasItxrInTopsColumn7(array $row): bool
{
    // Column 7 is index 6 in zero-based arrays.
    $topsName = trim((string) ($row[6] ?? ''));
    return $topsName !== '' && stripos($topsName, 'ITXR') !== false;
}

/**
 * Resolve the effective line-name key for a target row.
 * First tries exact column-7 match, then partial match against active keys.
 *
 * @param array<string, string> $effectiveMap
 */
function resolveEffectiveLineKey(string $column7Value, array $effectiveMap): string
{
    if ($column7Value === '') {
        return '';
    }

    if (array_key_exists($column7Value, $effectiveMap)) {
        return $column7Value;
    }

    foreach ($effectiveMap as $lineName => $_device) {
        if ($lineName !== '' && stripos($column7Value, $lineName) !== false) {
            return $lineName;
        }
    }

    return '';
}

/**
 * Apply replacements and row filtering for one round snapshot.
 *
 * @param array<int, array<int, string|null>> $targetRows
 * @param array<string, string> $effectiveMap
 * @return array<int, array<int, string|null>>
 */
function applyDeviceMap(array $targetRows, array $effectiveMap, string $roundLabel): array
{
    $resultRows = [];
    $totalRows = 0;
    $scopedRows = 0;
    $itxrRemovedRows = 0;
    $replacedRows = 0;

    foreach ($targetRows as $row) {
        $totalRows++;

        // Needs at least 7 columns to read line name and replace device name.
        if (!array_key_exists(6, $row) || !array_key_exists(5, $row)) {
            $resultRows[] = $row;
            continue;
        }

        $column7Value = trim((string) $row[6]);
        $resolvedLineKey = resolveEffectiveLineKey($column7Value, $effectiveMap);
        if ($resolvedLineKey === '') {
            $resultRows[] = $row;
            continue;
        }
        $scopedRows++;

        // Row is part of current cumulative move-round scope.
        // If its TOPS Name (column 7) partially matches ITXR, exclude it.
        if (rowHasItxrInTopsColumn7($row)) {
            $itxrRemovedRows++;
            continue;
        }

        $deviceName = trim((string) $row[5]);

        $srcDst = trim((string) ($row[3] ?? ''));
        $isSrc = strcasecmp($srcDst, 'SRC') === 0;
        $isItxe = stripos($deviceName, 'ITXE') !== false;
        if ($isSrc && $isItxe) {
            $row[5] = $effectiveMap[$resolvedLineKey];
            $replacedRows++;
        }

        $resultRows[] = $row;
    }

    debugEcho(
        "Round {$roundLabel}: total_rows={$totalRows}, scoped_rows={$scopedRows}, "
        . "itxr_removed={$itxrRemovedRows}, replaced={$replacedRows}, output_rows=" . count($resultRows)
    );

    return $resultRows;
}

/**
 * Write rows to CSV file.
 *
 * @param array<int, array<int, string|null>> $rows
 */
function writeCsvRows(string $path, array $rows): void
{
    $handle = fopen($path, 'wb');
    if ($handle === false) {
        throw new RuntimeException("Unable to write CSV: {$path}");
    }

    foreach ($rows as $row) {
        // Output only columns 1 through 6.
        $row = array_slice($row, 0, 6);
        // Explicitly pass $escape to avoid PHP 8.4+ deprecation warnings.
        fputcsv($handle, $row, ',', '"', '\\');
    }

    fclose($handle);
}

/**
 * Write per-interface output files for a given round.
 *
 * Pattern: YYYY-MM-DD-HH-MM-<interface_name>-<move round>.csv
 *
 * @param array<int, array<int, string|null>> $rows
 * @param array<int, string> $interfaceNames
 */
function writeInterfaceOutputs(
    array $rows,
    array $interfaceNames,
    string $outputDir,
    string $roundLabel,
    string $timestampPrefix
): void {
    foreach ($interfaceNames as $interfaceName) {
        $filteredRows = [];
        foreach ($rows as $row) {
            $rowInterfaceName = trim((string) ($row[0] ?? ''));
            if ($rowInterfaceName === $interfaceName) {
                $filteredRows[] = $row;
            }
        }

        $interfaceFilePart = sanitizeFilenamePart($interfaceName);
        $interfacePath = rtrim($outputDir, DIRECTORY_SEPARATOR)
            . DIRECTORY_SEPARATOR
            . "{$timestampPrefix}-{$interfaceFilePart}-{$roundLabel}.csv";
        writeCsvRows($interfacePath, $filteredRows);
        echo "Wrote: {$interfacePath}\n";
    }
}

/**
 * Write one total file and per-interface files for each move round.
 *
 * @param array<int, string> $roundKeys
 * @param array<int, array<int, string|null>> $targetRows
 * @param array<string, array{roundLabel: string, lineToDevice: array<string, string>}> $roundMap
 * @param array<int, string> $interfaceNames
 */
function processRounds(
    array $roundKeys,
    array $targetRows,
    array $roundMap,
    array $interfaceNames,
    string $outputDir,
    string $timestampPrefix,
    string $extension
): void {
    $effectiveMap = [];
    foreach ($roundKeys as $roundKeyRaw) {
        $roundKey = (string) $roundKeyRaw;
        if (!isset($roundMap[$roundKey])) {
            continue;
        }

        $roundLabel = $roundMap[$roundKey]['roundLabel'];
        foreach ($roundMap[$roundKey]['lineToDevice'] as $lineName => $deviceName) {
            $effectiveMap[$lineName] = $deviceName;
        }

        $updatedRows = applyDeviceMap($targetRows, $effectiveMap, $roundLabel);
        $outputPath = rtrim($outputDir, DIRECTORY_SEPARATOR)
            . DIRECTORY_SEPARATOR
            . "{$timestampPrefix}-total-{$roundLabel}{$extension}";
        writeCsvRows($outputPath, $updatedRows);

        echo "Wrote: {$outputPath}\n";
        writeInterfaceOutputs($updatedRows, $interfaceNames, $outputDir, $roundLabel, $timestampPrefix);
    }
}

try {
    $lookupRows = readCsvRows($lookupPath);
    $targetRows = readCsvRows($targetPath);
    $roundMap = buildRoundMap($lookupRows);

    if ($roundMap === []) {
        echo "No numeric move rounds found in lookup CSV column 9. No output files created.\n";
        exit(0);
    }

    $roundKeys = getSortedRoundKeys($roundMap);
    $interfaceNames = getUniqueInterfaceNames($targetRows);

    $targetInfo = pathinfo($targetPath);
    $dir = $targetInfo['dirname'] ?? '.';
    $extension = '.csv';
    $outputDir = $dir;

    $timestampPrefix = date('Y-m-d-H-i');

    processRounds($roundKeys, $targetRows, $roundMap, $interfaceNames, $outputDir, $timestampPrefix, $extension);
} catch (Throwable $e) {
    fwrite(STDERR, "Error: " . $e->getMessage() . "\n");
    exit(1);
}

