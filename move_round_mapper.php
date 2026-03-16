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
 * Build round map: [round => [lineName => deviceName]].
 *
 * @param array<int, array<int, string|null>> $lookupRows
 * @return array<int, array<string, string>>
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

        $round = (int) $roundRaw;
        if ($round < 2) {
            continue;
        }

        if (!isset($roundMap[$round])) {
            $roundMap[$round] = [];
        }

        // If duplicate line names exist in same round, later row wins.
        $roundMap[$round][$lineName] = $deviceName;
    }

    ksort($roundMap);
    return $roundMap;
}

/**
 * Apply replacements for rounds 2..$upToRound (cumulative).
 *
 * @param array<int, array<int, string|null>> $targetRows
 * @param array<int, array<string, string>> $roundMap
 * @return array<int, array<int, string|null>>
 */
function applyCumulativeRound(array $targetRows, array $roundMap, int $upToRound): array
{
    $effectiveMap = [];

    foreach ($roundMap as $round => $lineToDevice) {
        if ($round > $upToRound) {
            break;
        }

        foreach ($lineToDevice as $lineName => $deviceName) {
            $effectiveMap[$lineName] = $deviceName;
        }
    }

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
        if ($lineName !== '' && array_key_exists($lineName, $effectiveMap)) {
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
 * @param array<int, array<string, string>> $roundMap
 */
function processRoundsRecursively(
    int $currentRound,
    int $maxRound,
    array $targetRows,
    array $roundMap,
    string $outputBasePath,
    string $dateStamp,
    string $extension
): void {
    if ($currentRound > $maxRound) {
        return;
    }

    $updatedRows = applyCumulativeRound($targetRows, $roundMap, $currentRound);
    $outputPath = "{$outputBasePath}-{$currentRound}-{$dateStamp}{$extension}";
    writeCsvRows($outputPath, $updatedRows);

    echo "Wrote: {$outputPath}\n";

    processRoundsRecursively(
        $currentRound + 1,
        $maxRound,
        $targetRows,
        $roundMap,
        $outputBasePath,
        $dateStamp,
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

    $rounds = array_keys($roundMap);
    $maxRound = (int) max($rounds);

    $targetInfo = pathinfo($targetPath);
    $dir = $targetInfo['dirname'] ?? '.';
    $filename = $targetInfo['filename'] ?? 'output';
    $extension = isset($targetInfo['extension']) ? '.' . $targetInfo['extension'] : '';
    $outputBasePath = rtrim($dir, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . $filename;

    $dateStamp = date('Y-m-d');

    processRoundsRecursively(2, $maxRound, $targetRows, $roundMap, $outputBasePath, $dateStamp, $extension);
} catch (Throwable $e) {
    fwrite(STDERR, "Error: " . $e->getMessage() . "\n");
    exit(1);
}

