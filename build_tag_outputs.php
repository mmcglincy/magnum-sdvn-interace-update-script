<?php

declare(strict_types=1);

/**
 * Usage:
 *   php build_tag_outputs.php <lookup_csv> <tag_csv> [output_dir]
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
 * tag_csv columns:
 *   1) UUID
 *   2) NAME (TOPS)
 *   3) TAG1
 *   4) TAG2
 *   5) TAG3
 *   6) TAG4
 *   7) TAG5
 *   8) TAG6
 */

if ($argc < 3) {
    fwrite(STDERR, "Usage: php build_tag_outputs.php <lookup_csv> <tag_csv> [output_dir]\n");
    exit(1);
}

$lookupPath = $argv[1];
$tagPath = $argv[2];
$outputDir = $argv[3] ?? dirname($tagPath);

if (!is_file($lookupPath)) {
    fwrite(STDERR, "Lookup CSV not found: {$lookupPath}\n");
    exit(1);
}

if (!is_file($tagPath)) {
    fwrite(STDERR, "Tag CSV not found: {$tagPath}\n");
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
 * @param array<int, string|null> $row
 * @return array<int, string>
 */
function normalizeTagRow(array $row): array
{
    $normalized = [];
    for ($i = 0; $i < 8; $i++) {
        $normalized[$i] = (string) ($row[$i] ?? '');
    }

    return $normalized;
}

/**
 * Normalize move-round labels while allowing decimal points.
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
 * @param array<int, array<int, string|null>> $lookupRows
 * @return array<string, array<int, array{name: string, device: string, isNewTrue: bool}>>
 */
function buildLookupByRound(array $lookupRows): array
{
    $byRound = [];

    foreach ($lookupRows as $row) {
        // Name (3), Device (4), Move Round (9), New (12)
        if (!array_key_exists(2, $row) || !array_key_exists(3, $row) || !array_key_exists(8, $row)) {
            continue;
        }

        $name = trim((string) $row[2]);
        $device = trim((string) $row[3]);
        $roundLabel = normalizeRoundLabel((string) $row[8]);
        $newRaw = trim((string) ($row[11] ?? ''));
        $isNewTrue = strcasecmp($newRaw, 'TRUE') === 0;

        if ($name === '' || $device === '' || $roundLabel === '') {
            continue;
        }

        if (!isset($byRound[$roundLabel])) {
            $byRound[$roundLabel] = [];
        }

        $byRound[$roundLabel][] = [
            'name' => $name,
            'device' => $device,
            'isNewTrue' => $isNewTrue,
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
 * @param array<int, string> $row
 * @return array<int, string>
 */
function clearTransferredTags(array $row): array
{
    // TAG1..TAG6 are columns 3..8 (index 2..7)
    for ($i = 2; $i <= 7; $i++) {
        $row[$i] = '';
    }

    return $row;
}

/**
 * @param array<int, string> $row
 * @param array<int, string> $tags
 * @return array<int, string>
 */
function applyTransferredTags(array $row, array $tags): array
{
    for ($i = 2; $i <= 7; $i++) {
        $row[$i] = $tags[$i] ?? '';
    }

    return $row;
}

try {
    $lookupRows = readCsvRows($lookupPath);
    $tagRows = readCsvRows($tagPath);

    if ($tagRows === []) {
        throw new RuntimeException('Tag CSV is empty.');
    }

    $lookupByRound = buildLookupByRound($lookupRows);
    if ($lookupByRound === []) {
        echo "No valid move rounds found. No output files created.\n";
        exit(0);
    }

    $header = normalizeTagRow($tagRows[0]);
    $timestamp = date('Y-m-d');

    foreach ($lookupByRound as $roundLabel => $entries) {
        $outputRows = [$header];

        foreach ($entries as $entry) {
            if ($entry['isNewTrue']) {
                continue;
            }

            $name = $entry['name'];
            $perfectMatchIndex = -1;
            $itxrMatchIndex = -1;

            $itxrPattern = '/^' . preg_quote($name, '/') . '\s+ITXR$/i';

            foreach ($tagRows as $rowIndex => $tagRow) {
                if ($rowIndex === 0) {
                    continue;
                }

                $row = normalizeTagRow($tagRow);
                $topsName = trim($row[1]);
                if ($topsName === '') {
                    continue;
                }

                // Allow partial matches on NAME (TOPS).
                if (stripos($topsName, $name) === false) {
                    continue;
                }

                if ($perfectMatchIndex === -1 && strcasecmp($topsName, $name) === 0) {
                    $perfectMatchIndex = $rowIndex;
                }

                if ($itxrMatchIndex === -1 && preg_match($itxrPattern, $topsName) === 1) {
                    $itxrMatchIndex = $rowIndex;
                }
            }

            if ($perfectMatchIndex === -1) {
                continue;
            }

            $perfectRow = normalizeTagRow($tagRows[$perfectMatchIndex]);
            $capturedTags = $perfectRow;

            $outputRows[] = clearTransferredTags($perfectRow);

            if ($itxrMatchIndex !== -1) {
                $itxrRow = normalizeTagRow($tagRows[$itxrMatchIndex]);
                $outputRows[] = applyTransferredTags($itxrRow, $capturedTags);
            }
        }

        $outputPath = rtrim($outputDir, DIRECTORY_SEPARATOR)
            . DIRECTORY_SEPARATOR
            . "{$timestamp}-tag-{$roundLabel}.csv";
        writeCsvRows($outputPath, $outputRows);
        echo "Wrote: {$outputPath}\n";
    }
} catch (Throwable $e) {
    fwrite(STDERR, "Error: " . $e->getMessage() . "\n");
    exit(1);
}

