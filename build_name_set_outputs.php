<?php

declare(strict_types=1);

/**
 * Usage:
 *   php build_name_set_outputs.php <lookup_csv> <name_set_csv> [output_dir]
 *
 * lookup_csv columns:
 *   1) Schedule Time
 *   2) Schedule Move Dates
 *   3) Name
 *   4) Device
 *   5) Line Owner
 *   6) Information/Description
 *   7) SUB CT Assignment
 *   8) IP
 *   9) Move Round
 *   10) SDI
 *   11) Dual
 *   12) New
 *   13) Group to Act
 *   14) Person Responsible
 *   15) Move Complete - Need Sign-off
 *   16) Colmplete
 *   17) Status
 *   18) Notes
 *   19) Modified By
 *   20) Modified
 *
 * name_set_csv columns:
 *   1) Terminal ID
 *   2) Port Name
 *   3) Terminal Type
 *   4) Global
 *   5) Local
 *   6) TOPS
 *   7) VUE
 *   8) Description
 */

if ($argc < 3) {
    fwrite(STDERR, "Usage: php build_name_set_outputs.php <lookup_csv> <name_set_csv> [output_dir]\n");
    exit(1);
}

$lookupPath = $argv[1];
$nameSetPath = $argv[2];
$outputDir = $argv[3] ?? dirname($nameSetPath);

if (!is_file($lookupPath)) {
    fwrite(STDERR, "Lookup CSV not found: {$lookupPath}\n");
    exit(1);
}

if (!is_file($nameSetPath)) {
    fwrite(STDERR, "Name set CSV not found: {$nameSetPath}\n");
    exit(1);
}

if (!is_dir($outputDir)) {
    fwrite(STDERR, "Output directory not found: {$outputDir}\n");
    exit(1);
}

/**
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
 * @param array<int, array<int, string|null>> $rows
 */
function writeCsvRows(string $path, array $rows): void
{
    $handle = fopen($path, 'wb');
    if ($handle === false) {
        throw new RuntimeException("Unable to write CSV: {$path}");
    }

    foreach ($rows as $row) {
        fputcsv($handle, $row, ',', '"', '\\');
    }

    fclose($handle);
}

/**
 * Normalize round text for stable grouping (supports decimal points).
 */
function normalizeRoundLabel(string $roundRaw): string
{
    $value = trim($roundRaw);
    if ($value === '') {
        return '';
    }

    $value = preg_replace('/^\xEF\xBB\xBF/u', '', $value) ?? $value;
    $value = str_replace(',', '.', $value);
    if (!is_numeric($value)) {
        return '';
    }

    if (str_starts_with($value, '+')) {
        $value = substr($value, 1);
    }

    return $value === '-0' ? '0' : $value;
}

/**
 * Build grouped lookup rows by move round.
 *
 * @param array<int, array<int, string|null>> $lookupRows
 * @return array<string, array<int, array{name: string, device: string}>>
 */
function buildLookupByRound(array $lookupRows): array
{
    $byRound = [];

    foreach ($lookupRows as $row) {
        // Name (index 2), Device (index 3), Move Round (index 8)
        if (!array_key_exists(2, $row) || !array_key_exists(3, $row) || !array_key_exists(8, $row)) {
            continue;
        }

        $name = trim((string) $row[2]);
        $device = trim((string) $row[3]);
        $roundLabel = normalizeRoundLabel((string) $row[8]);

        if ($name === '' || $device === '' || $roundLabel === '') {
            continue;
        }

        if (!isset($byRound[$roundLabel])) {
            $byRound[$roundLabel] = [];
        }

        $byRound[$roundLabel][] = [
            'name' => $name,
            'device' => $device,
        ];
    }

    uksort(
        $byRound,
        static function (string $a, string $b): int {
            $cmp = (float) $a <=> (float) $b;
            return $cmp !== 0 ? $cmp : strcmp($a, $b);
        }
    );

    return $byRound;
}

/**
 * Ensure name-set rows have at least 8 columns.
 *
 * @param array<int, string|null> $row
 * @return array<int, string>
 */
function normalizeNameSetRow(array $row): array
{
    $normalized = [];
    for ($i = 0; $i < 8; $i++) {
        $normalized[$i] = (string) ($row[$i] ?? '');
    }

    return $normalized;
}

try {
    $lookupRows = readCsvRows($lookupPath);
    $nameSetRows = readCsvRows($nameSetPath);

    if ($nameSetRows === []) {
        throw new RuntimeException('Name set CSV is empty.');
    }

    $lookupByRound = buildLookupByRound($lookupRows);
    if ($lookupByRound === []) {
        echo "No valid move rounds found. No output files created.\n";
        exit(0);
    }

    $timestamp = date('Y-m-d');
    $headerRow = normalizeNameSetRow($nameSetRows[0]);

    foreach ($lookupByRound as $roundLabel => $entries) {
        $outputRows = [$headerRow];

        foreach ($entries as $entry) {
            $name = $entry['name'];
            $nameItxrPattern = '/^' . preg_quote($name, '/') . '-ITXR$/i';
            $replacementOld = 'z-old-' . $name;

            foreach ($nameSetRows as $rowIndex => $nameSetRow) {
                if ($rowIndex === 0) {
                    continue;
                }

                $row = normalizeNameSetRow($nameSetRow);
                $portName = $row[1];
                $tops = $row[5];

                // Partial match of lookup Name against TOPS.
                if (stripos($tops, $name) === false) {
                    continue;
                }

                if (preg_match($nameItxrPattern, trim($tops)) === 1) {
                    $row[4] = $name; // Local
                    $row[5] = $name; // TOPS
                    $row[6] = $name; // VUE
                    $outputRows[] = $row;
                    continue;
                }

                $portHasItxe = stripos($portName, 'ITXE') !== false;
                $portHasSrc1 = stripos($portName, 'SRC-1') !== false;
                if ($portHasItxe && $portHasSrc1) {
                    $row[4] = $replacementOld; // Local
                    $row[5] = $replacementOld; // TOPS
                    $row[6] = $replacementOld; // VUE
                    $outputRows[] = $row;
                }
            }
        }

        $outputPath = rtrim($outputDir, DIRECTORY_SEPARATOR)
            . DIRECTORY_SEPARATOR
            . "{$timestamp}-names-set-{$roundLabel}.csv";
        writeCsvRows($outputPath, $outputRows);
        echo "Wrote: {$outputPath}\n";
    }
} catch (Throwable $e) {
    fwrite(STDERR, "Error: " . $e->getMessage() . "\n");
    exit(1);
}

