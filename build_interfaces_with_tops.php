<?php

declare(strict_types=1);

/**
 * Usage:
 *   php build_interfaces_with_tops.php <edge_devices.xlsx> <interface_file> [output_dir]
 *
 * Output:
 *   YYYY-MM-DD-interfaces-all-with-tops.csv
 */

if ($argc < 3) {
    fwrite(STDERR, "Usage: php build_interfaces_with_tops.php <edge_devices.xlsx> <interface_file> [output_dir]\n");
    exit(1);
}

$xlsxPath = $argv[1];
$interfacePath = $argv[2];
$outputDir = $argv[3] ?? dirname($interfacePath);

if (!is_file($xlsxPath)) {
    fwrite(STDERR, "XLSX file not found: {$xlsxPath}\n");
    exit(1);
}

if (!is_file($interfacePath)) {
    fwrite(STDERR, "Interface file not found: {$interfacePath}\n");
    exit(1);
}

if (!is_dir($outputDir)) {
    fwrite(STDERR, "Output directory not found: {$outputDir}\n");
    exit(1);
}

if (!class_exists('ZipArchive')) {
    fwrite(STDERR, "ZipArchive extension is required to read XLSX files.\n");
    exit(1);
}

/**
 * Canonical interface mapping:
 * interface_name => [source_number, destination_number]
 *
 * @return array<string, array{source: string, destination: string}>
 */
function interfaceNumberMap(): array
{
    $raw = [
        'AFFILIATE FONT DSK' => ['4096', '4096'],
        'test_set' => ['4060', '4060'],
        'CT to STU TLM' => ['4055', '4055'],
        'CT to NOC TLM' => ['4084', '4084'],
        'DOTCOM DSK' => ['4097', '4097'],
        'QC LON' => ['4022', '4022'],
        'NEWSOURCE DSK' => ['4095', '4095'],
        'TOC to NOC TLM' => ['4080', '4080'],
        'TOC to STU TLM' => ['4056', '4056'],
        'TIMING' => ['4040', '4040'],
        'CT CORE_2' => ['4052', '4052'],
        'TIMING ACQ' => ['4041', '4041'],
        'CT DRE CH' => ['4073', '4073'],
        'JXS-TALLY' => ['4090', '4090'],
        'CLOUD-MASTER' => ['4085', '4085'],
    ];

    $map = [];
    foreach ($raw as $name => [$source, $destination]) {
        $map[normalizeInterfaceName($name)] = [
            'source' => $source,
            'destination' => $destination,
        ];
    }

    return $map;
}

function normalizeInterfaceName(string $name): string
{
    $name = trim($name);
    $name = preg_replace('/\s+/', ' ', $name) ?? $name;
    return strtolower($name);
}

function normalizeHeaderName(string $header): string
{
    $header = trim($header);
    $header = strtolower($header);
    return preg_replace('/[^a-z0-9]+/', '', $header) ?? $header;
}

function normalizeNumberKey(string $value): string
{
    $value = trim($value);
    if ($value === '') {
        return '';
    }

    $value = preg_replace('/^\xEF\xBB\xBF/u', '', $value) ?? $value;
    $value = str_replace(',', '.', $value);
    if (!is_numeric($value)) {
        return '';
    }

    if (str_contains($value, '.')) {
        $value = rtrim(rtrim($value, '0'), '.');
    }

    if (str_starts_with($value, '+')) {
        $value = substr($value, 1);
    }

    return $value === '-0' ? '0' : $value;
}

/**
 * @param array<int, string> $headerRow
 * @param array<int, string> $candidates
 */
function findHeaderIndex(array $headerRow, array $candidates): int
{
    $normalizedCandidates = [];
    foreach ($candidates as $candidate) {
        $normalizedCandidates[] = normalizeHeaderName($candidate);
    }

    // Exact normalized match first.
    foreach ($headerRow as $index => $headerName) {
        $normalizedHeader = normalizeHeaderName((string) $headerName);
        if (in_array($normalizedHeader, $normalizedCandidates, true)) {
            return (int) $index;
        }
    }

    // Fallback for sheets where headers include extra suffix/prefix text.
    foreach ($headerRow as $index => $headerName) {
        $normalizedHeader = normalizeHeaderName((string) $headerName);
        if ($normalizedHeader === '') {
            continue;
        }

        foreach ($normalizedCandidates as $candidate) {
            if (
                $candidate !== '' &&
                (str_contains($normalizedHeader, $candidate) || str_contains($candidate, $normalizedHeader))
            ) {
                return (int) $index;
            }
        }
    }

    return -1;
}

/**
 * Locate the actual header row in Edge Devices sheet and required columns.
 *
 * @param array<int, array<int, string>> $edgeRows
 * @return array{headerRowIndex: int, topsCol: int, srcCol: int, dstCol: int}
 */
function findEdgeHeaderAndColumns(array $edgeRows): array
{
    // Scan first N rows for the real header row.
    $scanLimit = min(count($edgeRows), 200);
    for ($rowIndex = 0; $rowIndex < $scanLimit; $rowIndex++) {
        $row = $edgeRows[$rowIndex];
        $topsCol = findHeaderIndex($row, ['Mnemonic;TOPS', 'Mnemonic TOPS', 'TOPS']);
        $srcCol = findHeaderIndex($row, ['Quartz SRCs', 'Quartz SRC']);
        $dstCol = findHeaderIndex($row, ['Quartz DSTs', 'Quartz DST']);
        if ($topsCol >= 0 && $srcCol >= 0 && $dstCol >= 0) {
            return [
                'headerRowIndex' => $rowIndex,
                'topsCol' => $topsCol,
                'srcCol' => $srcCol,
                'dstCol' => $dstCol,
            ];
        }
    }

    throw new RuntimeException('Required columns not found in Edge Devices sheet.');
}

/**
 * @return array<int, array<int, string>>
 */
function readDelimitedRows(string $path): array
{
    $handle = fopen($path, 'rb');
    if ($handle === false) {
        throw new RuntimeException("Unable to open interface file: {$path}");
    }

    $firstLine = fgets($handle);
    if ($firstLine === false) {
        fclose($handle);
        return [];
    }

    $delimiters = [",", "\t", ";", "|"];
    $bestDelimiter = ",";
    $bestCount = -1;
    foreach ($delimiters as $delimiter) {
        $count = count(str_getcsv($firstLine, $delimiter, '"', '\\'));
        if ($count > $bestCount) {
            $bestCount = $count;
            $bestDelimiter = $delimiter;
        }
    }

    rewind($handle);
    $rows = [];
    while (($row = fgetcsv($handle, 0, $bestDelimiter, '"', '\\')) !== false) {
        $normalized = [];
        foreach ($row as $value) {
            $normalized[] = (string) $value;
        }
        $rows[] = $normalized;
    }

    fclose($handle);
    return $rows;
}

/**
 * @param array<int, array<int, string>> $rows
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

function columnLettersToIndex(string $letters): int
{
    $letters = strtoupper($letters);
    $index = 0;
    for ($i = 0, $len = strlen($letters); $i < $len; $i++) {
        $index = ($index * 26) + (ord($letters[$i]) - ord('A') + 1);
    }

    return $index - 1;
}

/**
 * @return array<int, string>
 */
function readSharedStrings(ZipArchive $zip): array
{
    $xmlString = $zip->getFromName('xl/sharedStrings.xml');
    if ($xmlString === false) {
        return [];
    }

    $xml = simplexml_load_string($xmlString);
    if ($xml === false) {
        return [];
    }

    $strings = [];
    foreach ($xml->si as $si) {
        if (isset($si->t)) {
            $strings[] = (string) $si->t;
            continue;
        }

        $text = '';
        foreach ($si->r as $run) {
            $text .= (string) $run->t;
        }
        $strings[] = $text;
    }

    return $strings;
}

/**
 * @param array<int, string> $sharedStrings
 */
function cellValue(SimpleXMLElement $cell, array $sharedStrings): string
{
    $type = (string) $cell['t'];
    if ($type === 'inlineStr') {
        return isset($cell->is->t) ? (string) $cell->is->t : '';
    }

    $raw = isset($cell->v) ? (string) $cell->v : '';
    if ($type === 's') {
        $index = (int) $raw;
        return $sharedStrings[$index] ?? '';
    }

    return $raw;
}

/**
 * @return array<int, array<int, string>>
 */
function readXlsxSheetRows(string $xlsxPath, string $sheetName): array
{
    $zip = new ZipArchive();
    if ($zip->open($xlsxPath) !== true) {
        throw new RuntimeException("Unable to open XLSX file: {$xlsxPath}");
    }

    try {
        $workbookXmlString = $zip->getFromName('xl/workbook.xml');
        $relsXmlString = $zip->getFromName('xl/_rels/workbook.xml.rels');
        if ($workbookXmlString === false || $relsXmlString === false) {
            throw new RuntimeException('Invalid XLSX structure (missing workbook metadata).');
        }

        $workbookXml = simplexml_load_string($workbookXmlString);
        $relsXml = simplexml_load_string($relsXmlString);
        if ($workbookXml === false || $relsXml === false) {
            throw new RuntimeException('Unable to parse workbook metadata XML.');
        }

        $targetRelId = '';
        foreach ($workbookXml->sheets->sheet as $sheet) {
            if ((string) $sheet['name'] !== $sheetName) {
                continue;
            }

            $relAttrs = $sheet->attributes('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
            $targetRelId = (string) ($relAttrs['id'] ?? '');
            break;
        }

        if ($targetRelId === '') {
            throw new RuntimeException("Sheet not found: {$sheetName}");
        }

        $targetPath = '';
        foreach ($relsXml->Relationship as $relationship) {
            if ((string) $relationship['Id'] === $targetRelId) {
                $targetPath = (string) $relationship['Target'];
                break;
            }
        }

        if ($targetPath === '') {
            throw new RuntimeException("Sheet relationship not found for: {$sheetName}");
        }

        if (str_starts_with($targetPath, '/')) {
            $sheetPath = ltrim($targetPath, '/');
        } else {
            $sheetPath = 'xl/' . ltrim($targetPath, '/');
        }

        $sheetXmlString = $zip->getFromName($sheetPath);
        if ($sheetXmlString === false) {
            throw new RuntimeException("Unable to read sheet XML: {$sheetPath}");
        }

        $sharedStrings = readSharedStrings($zip);
        $sheetXml = simplexml_load_string($sheetXmlString);
        if ($sheetXml === false) {
            throw new RuntimeException('Unable to parse sheet XML.');
        }

        $rows = [];
        foreach ($sheetXml->sheetData->row as $rowNode) {
            $sparse = [];
            foreach ($rowNode->c as $cell) {
                $ref = (string) $cell['r'];
                if (preg_match('/^([A-Z]+)/i', $ref, $matches) !== 1) {
                    continue;
                }

                $colIndex = columnLettersToIndex($matches[1]);
                $sparse[$colIndex] = cellValue($cell, $sharedStrings);
            }

            // Keep sparse rows to avoid high memory usage on wide sheets.
            $rows[] = $sparse;
        }

        return $rows;
    } finally {
        $zip->close();
    }
}

try {
    $edgeRows = readXlsxSheetRows($xlsxPath, 'Edge Devices');
    if ($edgeRows === []) {
        throw new RuntimeException('Edge Devices sheet is empty.');
    }

    $edgeHeaderInfo = findEdgeHeaderAndColumns($edgeRows);
    $topsCol = $edgeHeaderInfo['topsCol'];
    $srcCol = $edgeHeaderInfo['srcCol'];
    $dstCol = $edgeHeaderInfo['dstCol'];
    $edgeHeaderRowIndex = $edgeHeaderInfo['headerRowIndex'];

    $topsBySrcNumber = [];
    $topsByDstNumber = [];
    foreach ($edgeRows as $rowIndex => $row) {
        if ($rowIndex <= $edgeHeaderRowIndex) {
            continue;
        }

        $tops = trim((string) ($row[$topsCol] ?? ''));
        if ($tops === '') {
            continue;
        }

        $srcKey = normalizeNumberKey((string) ($row[$srcCol] ?? ''));
        $dstKey = normalizeNumberKey((string) ($row[$dstCol] ?? ''));
        if ($srcKey !== '' && !isset($topsBySrcNumber[$srcKey])) {
            $topsBySrcNumber[$srcKey] = $tops;
        }
        if ($dstKey !== '' && !isset($topsByDstNumber[$dstKey])) {
            $topsByDstNumber[$dstKey] = $tops;
        }
    }

    $interfaceRows = readDelimitedRows($interfacePath);
    if ($interfaceRows === []) {
        throw new RuntimeException('Interface file is empty.');
    }

    $interfaceHeader = $interfaceRows[0];
    $interfaceNameCol = findHeaderIndex($interfaceHeader, ['Interface Name']);
    $srcDstCol = findHeaderIndex($interfaceHeader, ['SRC/DST']);
    $orderCol = findHeaderIndex($interfaceHeader, ['Order']);
    if ($interfaceNameCol < 0 || $srcDstCol < 0 || $orderCol < 0) {
        throw new RuntimeException('Required columns not found in interface file.');
    }

    $interfaceMap = interfaceNumberMap();

    $outputRows = [];
    $headerOut = $interfaceHeader;
    $headerOut[] = 'TOPS Name';
    $outputRows[] = $headerOut;

    foreach ($interfaceRows as $rowIndex => $row) {
        if ($rowIndex === 0) {
            continue;
        }

        $interfaceName = trim((string) ($row[$interfaceNameCol] ?? ''));
        $srcDst = strtoupper(trim((string) ($row[$srcDstCol] ?? '')));
        $orderKey = normalizeNumberKey((string) ($row[$orderCol] ?? ''));

        $topsName = '';
        $interfaceKey = normalizeInterfaceName($interfaceName);
        if (isset($interfaceMap[$interfaceKey]) && $orderKey !== '') {
            $expected = $srcDst === 'SRC'
                ? $interfaceMap[$interfaceKey]['source']
                : $interfaceMap[$interfaceKey]['destination'];

            if ($orderKey === $expected) {
                if ($srcDst === 'SRC') {
                    $topsName = $topsBySrcNumber[$orderKey] ?? '';
                } elseif ($srcDst === 'DST') {
                    $topsName = $topsByDstNumber[$orderKey] ?? '';
                }
            }
        }

        $rowOut = $row;
        $rowOut[] = $topsName;
        $outputRows[] = $rowOut;
    }

    $dateStamp = date('Y-m-d');
    $outputPath = rtrim($outputDir, DIRECTORY_SEPARATOR)
        . DIRECTORY_SEPARATOR
        . "{$dateStamp}-interfaces-all-with-tops.csv";

    writeCsvRows($outputPath, $outputRows);
    echo "Wrote: {$outputPath}\n";
} catch (Throwable $e) {
    fwrite(STDERR, "Error: " . $e->getMessage() . "\n");
    exit(1);
}

