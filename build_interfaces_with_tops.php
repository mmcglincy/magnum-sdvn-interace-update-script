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

function normalizeInterfaceName(string $name): string
{
    $name = trim($name);
    $name = preg_replace('/[^a-z0-9]+/i', ' ', $name) ?? $name;
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
    if (preg_match('/[-+]?\d+(?:[.,]\d+)?/', $value, $matches) !== 1) {
        return '';
    }

    $value = str_replace(',', '.', $matches[0]);
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

function debugEcho(string $message): void
{
    echo "[DEBUG] {$message}\n";
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
 * @return array{headerRowIndex: int, topsCol: int}
 */
function findEdgeHeaderAndColumns(array $edgeRows): array
{
    // Scan first N rows for the real header row.
    $scanLimit = min(count($edgeRows), 200);
    for ($rowIndex = 0; $rowIndex < $scanLimit; $rowIndex++) {
        $row = $edgeRows[$rowIndex];
        $topsCol = findHeaderIndex($row, ['Mnemonic;TOPS', 'Mnemonic TOPS', 'TOPS']);
        $hasQuartzColumns = false;
        foreach ($row as $headerCell) {
            $headerNorm = normalizeHeaderName((string) $headerCell);
            if (str_contains($headerNorm, 'quartzsrcs') || str_contains($headerNorm, 'quartzdsts')) {
                $hasQuartzColumns = true;
                break;
            }
        }

        if ($topsCol >= 0 && $hasQuartzColumns) {
            return [
                'headerRowIndex' => $rowIndex,
                'topsCol' => $topsCol,
            ];
        }
    }

    throw new RuntimeException('Required columns not found in Edge Devices sheet.');
}

/**
 * @return array{direction: string, listen: string, interface: string}|null
 */
function parseEdgeInterfaceHeader(string $headerCell): ?array
{
    $raw = trim($headerCell);
    if ($raw === '') {
        return null;
    }

    $parts = array_map('trim', explode(';', $raw));
    if (count($parts) < 3) {
        return null;
    }

    $directionRaw = strtolower($parts[0]);
    $direction = '';
    if (str_contains($directionRaw, 'quartz src')) {
        $direction = 'SRC';
    } elseif (str_contains($directionRaw, 'quartz dst')) {
        $direction = 'DST';
    }
    if ($direction === '') {
        return null;
    }

    $listen = normalizeNumberKey($parts[1]);
    $interface = normalizeInterfaceName(implode(';', array_slice($parts, 2)));
    if ($listen === '' || $interface === '') {
        return null;
    }

    return [
        'direction' => $direction,
        'listen' => $listen,
        'interface' => $interface,
    ];
}

/**
 * Build map from interface + listen + direction to column index.
 *
 * @param array<int, string> $edgeHeaderRow
 * @return array<string, array<string, array<string, int>>>
 */
function buildEdgeColumnLookup(array $edgeHeaderRow): array
{
    $lookup = [
        'SRC' => [],
        'DST' => [],
    ];

    foreach ($edgeHeaderRow as $colIndex => $headerCell) {
        $parsed = parseEdgeInterfaceHeader((string) $headerCell);
        if ($parsed === null) {
            continue;
        }

        $direction = $parsed['direction'];
        $interface = $parsed['interface'];
        $listen = $parsed['listen'];
        if (!isset($lookup[$direction][$interface])) {
            $lookup[$direction][$interface] = [];
        }

        $lookup[$direction][$interface][$listen] = (int) $colIndex;
    }

    return $lookup;
}

/**
 * Build TOPS lookup by direction/interface/listen/order from edge data rows.
 *
 * @param array<int, array<int, string>> $edgeRows
 * @param array<string, array<string, array<string, int>>> $edgeColumnLookup
 * @return array<string, array<string, array<string, array<string, string>>>>
 */
function buildTopsLookup(
    array $edgeRows,
    int $edgeHeaderRowIndex,
    int $topsCol,
    array $edgeColumnLookup
): array {
    $topsByDirectionInterfaceListenOrder = [
        'SRC' => [],
        'DST' => [],
    ];

    foreach ($edgeRows as $rowIndex => $row) {
        if ($rowIndex <= $edgeHeaderRowIndex) {
            continue;
        }

        $tops = trim((string) ($row[$topsCol] ?? ''));
        if ($tops === '') {
            continue;
        }

        foreach (['SRC', 'DST'] as $direction) {
            foreach ($edgeColumnLookup[$direction] as $interfaceKey => $listenToCol) {
                foreach ($listenToCol as $listenKey => $colIndex) {
                    $orderKey = normalizeNumberKey((string) ($row[$colIndex] ?? ''));
                    if ($orderKey === '') {
                        continue;
                    }

                    if (!isset($topsByDirectionInterfaceListenOrder[$direction][$interfaceKey])) {
                        $topsByDirectionInterfaceListenOrder[$direction][$interfaceKey] = [];
                    }
                    if (!isset($topsByDirectionInterfaceListenOrder[$direction][$interfaceKey][$listenKey])) {
                        $topsByDirectionInterfaceListenOrder[$direction][$interfaceKey][$listenKey] = [];
                    }
                    if (!isset($topsByDirectionInterfaceListenOrder[$direction][$interfaceKey][$listenKey][$orderKey])) {
                        $topsByDirectionInterfaceListenOrder[$direction][$interfaceKey][$listenKey][$orderKey] = $tops;
                    }
                }
            }
        }
    }

    return $topsByDirectionInterfaceListenOrder;
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
    $edgeHeaderRowIndex = $edgeHeaderInfo['headerRowIndex'];
    debugEcho("Edge header row={$edgeHeaderRowIndex}, topsCol={$topsCol}");

    $edgeHeaderRow = $edgeRows[$edgeHeaderRowIndex];
    $edgeColumnLookup = buildEdgeColumnLookup($edgeHeaderRow);
    $topsLookup = buildTopsLookup($edgeRows, $edgeHeaderRowIndex, $topsCol, $edgeColumnLookup);
    debugEcho('Edge SRC interface columns=' . count($edgeColumnLookup['SRC']));
    debugEcho('Edge DST interface columns=' . count($edgeColumnLookup['DST']));
    debugEcho('TOPS SRC interfaces in lookup=' . count($topsLookup['SRC']));
    debugEcho('TOPS DST interfaces in lookup=' . count($topsLookup['DST']));

    $interfaceRows = readDelimitedRows($interfacePath);
    if ($interfaceRows === []) {
        throw new RuntimeException('Interface file is empty.');
    }

    $interfaceHeader = $interfaceRows[0];
    if (count($interfaceHeader) < 5) {
        throw new RuntimeException('Interface file must contain at least 5 columns.');
    }

    // Use fixed column positions from the interface file specification:
    // 1) Interface Name, 4) SRC/DST, 5) Order.
    $interfaceNameCol = 0;
    $srcDstCol = 3;
    $orderCol = 4;
    debugEcho('Using fixed interface columns: Interface Name=1, SRC/DST=4, Order=5');

    $outputRows = [];
    $headerOut = $interfaceHeader;
    $headerOut[] = 'TOPS Name';
    $outputRows[] = $headerOut;

    $totalRows = 0;
    $matchedRows = 0;
    $unmappedInterfaceRows = 0;
    $emptyOrderRows = 0;
    $missingEdgeColumnRows = 0;
    $noTopsRows = 0;
    $rowDebugLimit = 100;
    $rowDebugCount = 0;

    foreach ($interfaceRows as $rowIndex => $row) {
        if ($rowIndex === 0) {
            continue;
        }
        $totalRows++;

        $interfaceName = trim((string) ($row[$interfaceNameCol] ?? ''));
        $listenPortKey = normalizeNumberKey((string) ($row[1] ?? ''));
        $srcDst = strtoupper(trim((string) ($row[$srcDstCol] ?? '')));
        $orderKey = normalizeNumberKey((string) ($row[$orderCol] ?? ''));

        $topsName = '';
        $interfaceKey = normalizeInterfaceName($interfaceName);
        $direction = $srcDst === 'DST' ? 'DST' : 'SRC';
        if (!isset($edgeColumnLookup[$direction][$interfaceKey])) {
            $unmappedInterfaceRows++;
            if ($rowDebugCount < $rowDebugLimit) {
                debugEcho("Row {$rowIndex}: interface not found in Edge columns: '{$interfaceName}' ({$direction})");
                $rowDebugCount++;
            }
        } elseif ($listenPortKey === '') {
            $missingEdgeColumnRows++;
            if ($rowDebugCount < $rowDebugLimit) {
                debugEcho("Row {$rowIndex}: empty/non-numeric Listen Port for '{$interfaceName}'");
                $rowDebugCount++;
            }
        } elseif ($orderKey === '') {
            $emptyOrderRows++;
            if ($rowDebugCount < $rowDebugLimit) {
                debugEcho("Row {$rowIndex}: empty/non-numeric Order for '{$interfaceName}'");
                $rowDebugCount++;
            }
        } else {
            $listenColumns = $edgeColumnLookup[$direction][$interfaceKey];
            $targetListenKey = $listenPortKey;
            if (!isset($listenColumns[$targetListenKey])) {
                $missingEdgeColumnRows++;
                if ($rowDebugCount < $rowDebugLimit) {
                    debugEcho(
                        "Row {$rowIndex}: no Edge {$direction} column for interface='{$interfaceName}', Listen Port='{$listenPortKey}'"
                    );
                    $rowDebugCount++;
                }
            } else {
                $topsName = $topsLookup[$direction][$interfaceKey][$targetListenKey][$orderKey] ?? '';
            }
        }

        if ($topsName === '') {
            $noTopsRows++;
            if ($rowDebugCount < $rowDebugLimit) {
                debugEcho(
                    "Row {$rowIndex}: no TOPS match (Interface='{$interfaceName}', SRC/DST='{$srcDst}', Order='{$orderKey}')"
                );
                $rowDebugCount++;
            }
        } else {
            $matchedRows++;
        }

        $rowOut = $row;
        $rowOut[] = $topsName;
        $outputRows[] = $rowOut;
    }
    debugEcho("Processed interface rows={$totalRows}");
    debugEcho("Matched TOPS rows={$matchedRows}");
    debugEcho("No TOPS rows={$noTopsRows}");
    debugEcho("Interface not found in Edge columns rows={$unmappedInterfaceRows}");
    debugEcho("Missing/invalid Listen Port or Edge column rows={$missingEdgeColumnRows}");
    debugEcho("Empty/non-numeric Order rows={$emptyOrderRows}");

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

