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

    while (($row = fgetcsv($handle)) !== false) {
        $rows[] = $row;
    }

    fclose($handle);
    return $rows;
}

/**
 * Build round map: [roundLabel => [lineName => deviceName]].
 *
 * - round labels can be decimals (example: 2.5)
 * - rows without a specified move round are skipped
 *
 * @param array<int, array<int, string|null>> $lookupRows
 * @return array<string, array<string, string>>
 */
function buildRoundMap(array $lookupRows): array
{
    $roundMap = [];

    foreach ($lookupRows as $row) {
        // Column 3 (index 2): line name
        // Column 4 (index 3): device name
        // Column 9 (index 8): move round
        if (!array_key_exists(8, $row) || !array_key_exists(2, $row) || !array_key_exists(3, $row)) {
            continue;
        }

        $lineName = trim((string) $row[2]);
        $deviceName = (string) $row[3];
        $roundRaw = trim((string) $row[8]);

        if ($lineName === '' || $roundRaw === '' || !is_numeric($roundRaw)) {
            continue;
        }

        $roundValue = (float) $roundRaw;
        if ($roundValue < 2.0) {
            continue;
        }

        if (!isset($roundMap[$roundRaw])) {
            $roundMap[$roundRaw] = [];
        }

        // If duplicate line names exist in same round, later row wins.
        $roundMap[$roundRaw][$lineName] = $deviceName;
    }

    return $roundMap;
}

/**
 * Get round labels sorted by numeric value (ascending).
 *
 * @param array<string, array<string, string>> $roundMap
 * @return array<int, string>
 */
function getSortedRoundLabels(array $roundMap): array
{
    $roundLabels = array_keys($roundMap);
    usort(
        $roundLabels,
        static function (string $a, string $b): int {
            $valueCompare = (float) $a <=> (float) $b;
            if ($valueCompare !== 0) {
                return $valueCompare;
            }

            // Deterministic tie-breaker if numeric values are equal.
            return strcmp($a, $b);
        }
    );

    return $roundLabels;
}

/**
 * Apply replacements from a prepared line->device map.
 *
 * Rules:
 * - If a target row line-name matches an active move-round mapping and device is ITXR, remove the row.
 * - Only replace when a target row line-name matches, column 4 is SRC, and existing device is ITXE.
 *
 * @param array<int, array<int, string|null>> $targetRows
 * @param array<string, string> $effectiveMap
 * @return array<int, array<int, string|null>>
 */
function applyDeviceMap(array $targetRows, array $effectiveMap): array
{
    if ($effectiveMap === []) {
        return $targetRows;
    }

    $resultRows = [];
    foreach ($targetRows as $row) {
        // Needs at least 7 columns to read line name and replace device name.
        if (!array_key_exists(6, $row) || !array_key_exists(5, $row)) {
            $resultRows[] = $row;
            continue;
        }

        $lineName = trim((string) $row[6]);
        if ($lineName === '' || !array_key_exists($lineName, $effectiveMap)) {
            $resultRows[] = $row;
            continue;
        }

        $deviceName = trim((string) $row[5]);
        if (stripos($deviceName, 'ITXR') !== false) {
            // If already ITXR for a matched line, drop this row.
            continue;
        }

        $srcDst = trim((string) ($row[3] ?? ''));
        $isSrc = strcasecmp($srcDst, 'SRC') === 0;
        $isItxe = stripos($deviceName, 'ITXE') !== false;
        if ($isSrc && $isItxe) {
            $row[5] = $effectiveMap[$lineName];
        }

        $resultRows[] = $row;
    }

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
        fputcsv($handle, $row);
    }

    fclose($handle);
}

/**
 * Recursively write one output file per round.
 *
 * @param array<int, array<int, string|null>> $targetRows
 * @param array<int, string> $roundLabels
 * @param array<string, array<string, string>> $roundMap
 * @param array<string, string> $effectiveMap
 */
function processRoundsRecursively(
    int $roundIndex,
    array $roundLabels,
    array $targetRows,
    array $roundMap,
    array $effectiveMap,
    string $outputDir,
    string $timestampPrefix,
    string $extension
): void {
    if (!isset($roundLabels[$roundIndex])) {
        return;
    }

    $roundLabel = $roundLabels[$roundIndex];
    foreach ($roundMap[$roundLabel] as $lineName => $deviceName) {
        $effectiveMap[$lineName] = $deviceName;
    }

    $updatedRows = applyDeviceMap($targetRows, $effectiveMap);
    $outputPath = rtrim($outputDir, DIRECTORY_SEPARATOR)
        . DIRECTORY_SEPARATOR
        . "{$timestampPrefix}-total-{$roundLabel}{$extension}";
    writeCsvRows($outputPath, $updatedRows);

    echo "Wrote: {$outputPath}\n";

    processRoundsRecursively(
        $roundIndex + 1,
        $roundLabels,
        $targetRows,
        $roundMap,
        $effectiveMap,
        $outputDir,
        $timestampPrefix,
        $extension
    );
}

try {
    $lookupRows = readCsvRows($lookupPath);
    $targetRows = readCsvRows($targetPath);
    $roundMap = buildRoundMap($lookupRows);

    if ($roundMap === []) {
        echo "No move rounds >= 2 found in lookup CSV. No output files created.\n";
        exit(0);
    }

    $roundLabels = getSortedRoundLabels($roundMap);

    $targetInfo = pathinfo($targetPath);
    $dir = $targetInfo['dirname'] ?? '.';
    $extension = isset($targetInfo['extension']) ? '.' . $targetInfo['extension'] : '';
    $outputDir = $dir;

    $timestampPrefix = date('Y-m-d-H-i');

    processRoundsRecursively(0, $roundLabels, $targetRows, $roundMap, [], $outputDir, $timestampPrefix, $extension);
} catch (Throwable $e) {
    fwrite(STDERR, "Error: " . $e->getMessage() . "\n");
    exit(1);
}

