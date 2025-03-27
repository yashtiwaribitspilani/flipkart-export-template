<?php

declare(strict_types=1);

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use Exception;

/**
 * List of countries in the world.
 */
const COUNTRIES = [
    'Afghanistan',
    'Albania',
    'Algeria',
    'Andorra',
    'Angola',
    'Antigua and Barbuda',
    'Argentina',
    'Armenia',
    'Australia',
    'Austria',
    'Azerbaijan',
    'Bahamas',
    'Bahrain',
    'Bangladesh',
    'Barbados',
    'Belarus',
    'Belgium',
    'Belize',
    'Benin',
    'Bhutan',
    'Bolivia',
    'Bosnia and Herzegovina',
    'Botswana',
    'Brazil',
    'Brunei',
    'Bulgaria',
    'Burkina Faso',
    'Burundi',
    'Cabo Verde',
    'Cambodia',
    'Cameroon',
    'Canada',
    'Central African Republic',
    'Chad',
    'Chile',
    'China',
    'Colombia',
    'Comoros',
    'Congo (Congo-Brazzaville)',
    'Costa Rica',
    'Croatia',
    'Cuba',
    'Cyprus',
    'Czechia',
    'Democratic Republic of the Congo',
    'Denmark',
    'Djibouti',
    'Dominica',
    'Dominican Republic',
    'Ecuador',
    'Egypt',
    'El Salvador',
    'Equatorial Guinea',
    'Eritrea',
    'Estonia',
    'Eswatini',
    'Ethiopia',
    'Fiji',
    'Finland',
    'France',
    'Gabon',
    'Gambia',
    'Georgia',
    'Germany',
    'Ghana',
    'Greece',
    'Grenada',
    'Guatemala',
    'Guinea',
    'Guinea-Bissau',
    'Guyana',
    'Haiti',
    'Holy See',
    'Honduras',
    'Hungary',
    'Iceland',
    'India',
    'Indonesia',
    'Iran',
    'Iraq',
    'Ireland',
    'Israel',
    'Italy',
    'Jamaica',
    'Japan',
    'Jordan',
    'Kazakhstan',
    'Kenya',
    'Kiribati',
    'Kuwait',
    'Kyrgyzstan',
    'Laos',
    'Latvia',
    'Lebanon',
    'Lesotho',
    'Liberia',
    'Libya',
    'Liechtenstein',
    'Lithuania',
    'Luxembourg',
    'Madagascar',
    'Malawi',
    'Malaysia',
    'Maldives',
    'Mali',
    'Malta',
    'Marshall Islands',
    'Mauritania',
    'Mauritius',
    'Mexico',
    'Micronesia',
    'Moldova',
    'Monaco',
    'Mongolia',
    'Montenegro',
    'Morocco',
    'Mozambique',
    'Myanmar (Burma)',
    'Namibia',
    'Nauru',
    'Nepal',
    'Netherlands',
    'New Zealand',
    'Nicaragua',
    'Niger',
    'Nigeria',
    'North Korea',
    'North Macedonia',
    'Norway',
    'Oman',
    'Pakistan',
    'Palau',
    'Palestine State',
    'Panama',
    'Papua New Guinea',
    'Paraguay',
    'Peru',
    'Philippines',
    'Poland',
    'Portugal',
    'Qatar',
    'Romania',
    'Russia',
    'Rwanda',
    'Saint Kitts and Nevis',
    'Saint Lucia',
    'Saint Vincent and the Grenadines',
    'Samoa',
    'San Marino',
    'Sao Tome and Principe',
    'Saudi Arabia',
    'Senegal',
    'Serbia',
    'Seychelles',
    'Sierra Leone',
    'Singapore',
    'Slovakia',
    'Slovenia',
    'Solomon Islands',
    'Somalia',
    'South Africa',
    'South Korea',
    'South Sudan',
    'Spain',
    'Sri Lanka',
    'Sudan',
    'Suriname',
    'Sweden',
    'Switzerland',
    'Syria',
    'Tajikistan',
    'Tanzania',
    'Thailand',
    'Timor-Leste',
    'Togo',
    'Tonga',
    'Trinidad and Tobago',
    'Tunisia',
    'Turkey',
    'Turkmenistan',
    'Tuvalu',
    'Uganda',
    'Ukraine',
    'United Arab Emirates',
    'United Kingdom',
    'United States of America',
    'Uruguay',
    'Uzbekistan',
    'Vanuatu',
    'Venezuela',
    'Vietnam',
    'Yemen',
    'Zambia',
    'Zimbabwe'
];

/**
 * Constant for the "Fullfilment by" field.
 */
const FULLFILMENT_BY_VALUE = 'Seller';

/**
 * Constant for the "Shipping provider" field.
 */
const SHIPPING_PROVIDER_VALUE = 'Flipkart';

/**
 * Checks if a given value is a positive integer.
 *
 * @param mixed $value
 * @return bool
 */
function IsPositiveInteger($value): bool
{
    if (!is_numeric($value)) {
        return false;
    }
    $num = floatval($value);
    return ($num > 0 && floor($num) === $num);
}

/**
 * Checks if a value is one of the allowed length units.
 *
 * @param mixed $value
 * @return bool
 */
function IsValidLengthUnit($value): bool
{
    $allowedUnits = ['cm', 'mm', 'inch'];
    return in_array((string)$value, $allowedUnits, true);
}

/**
 * Normalizes the given country value.
 * Performs a case-insensitive match against COUNTRIES and returns the standardized version.
 *
 * @param mixed $value
 * @return bool|string Returns the normalized country name if found, otherwise false.
 */
function normalizeCountry($value)
{
    $input = strtolower(trim((string)$value));
    foreach (COUNTRIES as $country) {
        if (strtolower($country) === $input) {
            return $country;
        }
    }
    return false;
}

/**
 * Checks if a value is a valid URL.
 *
 * @param mixed $value
 * @return bool
 */
function IsValidUrl($value): bool
{
    $value = trim((string)$value);
    return filter_var($value, FILTER_VALIDATE_URL) !== false;
}

/**
 * Main execution function.
 * Contains all execution logic.
 *
 * @return void
 */
function main(): void
{
    // --- Allowed sets for specific columns ---
    $allowedTaxCodes = ['GST_0', 'GST_12', 'GST_18', 'GST_3', 'GST_5', 'GST_APPAREL'];
    $allowedBrands = [
        'Beige', 'Black', 'Blue', 'Brown', 'Clear', 'Gold', 'Green', 'Grey',
        'Khaki', 'Maroon', 'Multicolor', 'Orange', 'Pink', 'Purple', 'Red',
        'Silver', 'Tan', 'White', 'Yellow'
    ];
    $allowedTypes = [
        'Clutch', 'Hand-held Bag', 'Hobo', 'Messenger Bag', 'Satchel',
        'Shoulder Bag', 'Sling Bag', 'Tote'
    ];
    $allowedIdealFor = ['Boys', 'Boys & Girls', 'Girls', 'Men', 'Men & Women', 'Women'];
    $allowedOccasions = ['Casual', 'Evening/Party', 'Formal', 'Sports'];
    $allowedMaterials = [
        'Acrylic', 'Beads', 'Brocade', 'Canvas', 'Cotton', 'Denim', 'Fabric',
        'Flex', 'Genuine Leather', 'Juco', 'Jute', 'Leatherette', 'Metal',
        'Natural Fibre', 'PU', 'Plastic', 'Polyester', 'Rexine', 'Satin',
        'Silicon', 'Silk', 'Synthetic Leather', 'Tyvek', 'Velvet', 'Wood', 'Wool'
    ];

    // Define variables for required flag.
    $isRequired = true;
    $isNotRequired = false;

    // --- Step 1: Read the CSV data (using header names) ---
    $baseInputFilePath = 'sample_data.csv'; // Replace with your CSV file path

    try {
        $csvReader = IOFactory::createReader('Csv');
        $csvReader->setReadDataOnly(true);
        $baseInputSpreadsheet = $csvReader->load($baseInputFilePath);
    } catch (Exception $e) {
        error_log("Error loading CSV file: " . $e->getMessage());
        echo "Error: Unable to load CSV file.\n";
        return;
    }

    $basebaseInputWorksheet = $baseInputSpreadsheet->getActiveSheet();
    $rows = $basebaseInputWorksheet->toArray(null, true, true, true);

    // Assume the first row contains headers.
    $baseInputHeaders = array_map('trim', array_values(reset($rows)));
    array_shift($rows);

    // Convert each row to an associative array using header names.
    $inputData = [];
    foreach ($rows as $row) {
        $assoc = [];
        $colIndex = 0;
        foreach ($baseInputHeaders as $header) {
            $letter = Coordinate::stringFromColumnIndex($colIndex + 1);
            $assoc[$header] = isset($row[$letter]) ? $row[$letter] : null;
            $colIndex++;
        }
        $inputData[] = $assoc;
    }

    // --- Step 2: Open the Flipkart template ---
    $flipkartTemplatePath = 'C_sling-bag_fd927b15e6244645_1703-2438FK_REQH2ILIQXHAH.xlsx';

    try {
        $flipkartTemplateSpreadsheet = IOFactory::load($flipkartTemplatePath);
    } catch (Exception $e) {
        error_log("Error loading template file: " . $e->getMessage());
        echo "Error: Unable to load template file.\n";
        return;
    }

    $flipkartSheet = $flipkartTemplateSpreadsheet->getSheetByName('sling_bag');

    // Build a lookup for template headers (assumes headers are in row 1 and match CSV header names).
    $templateHeaders = array_map('trim', $flipkartSheet->rangeToArray('A1:' . $flipkartSheet->getHighestColumn() . '1')[0]);
    $templateHeaderMap = [];
    foreach ($templateHeaders as $index => $headerValue) {
        $templateHeaderMap[$headerValue] = Coordinate::stringFromColumnIndex($index + 1);
    }

    // --- Step 3: Define the mapping ---
    // Mapping: target header name => [is_required, source CSV header title]
    $mapping = [
        'Seller SKU ID'                => [$isRequired, 'Seller SKU ID'],
        'MRP (INR)'                    => [$isRequired, 'MRP (INR)'],
        'Your selling price (INR)'     => [$isRequired, 'Your selling price (INR)'],
        'Fullfilment by'               => [$isRequired, 'Fullfilment by'],
        'Procurement SLA (DAY)'        => [$isRequired, 'Procurement SLA (DAY)'],
        'Stock'                        => [$isRequired, 'Stock'],
        'Shipping provider'            => [$isRequired, 'Shipping provider'],
        'Local delivery charge (INR)'  => [$isRequired, 'Local delivery charge (INR)'],
        'Zonal delivery charge (INR)'  => [$isRequired, 'Zonal delivery charge (INR)'],
        'National delivery charge (INR)' => [$isRequired, 'National delivery charge (INR)'],
        'Height (CM)'                  => [$isRequired, 'Height (CM)'],
        'Weight (KG)'                  => [$isRequired, 'Weight (KG)'],
        'Breadth (CM)'                 => [$isRequired, 'Breadth (CM)'],
        'Length (CM)'                  => [$isRequired, 'Length (CM)'],
        'Country Of Origin'            => [$isRequired, 'Country Of Origin'],
        'Manufacturer Details'         => [$isRequired, 'Manufacturer Details'],
        'Packer Details'               => [$isRequired, 'Packer Details'],
        'Tax Code'                     => [$isRequired, 'Tax Code'],
        'Brand'                        => [$isRequired, 'Brand'],
        'Model Name'                   => [$isRequired, 'Model Name'],
        'Brand Color'                  => [$isRequired, 'Brand Color'],
        'Color'                        => [$isRequired, 'Color'],
        'Style Code'                   => [$isRequired, 'Style Code'],
        'Type'                         => [$isRequired, 'Type'],
        'Ideal For'                    => [$isRequired, 'Ideal For'],
        'Occasion'                     => [$isRequired, 'Occasion'],
        'Material'                     => [$isRequired, 'Material'],
        'Pack of'                      => [$isRequired, 'Pack of'],
        'Height'                       => [$isRequired, 'Height'],
        'Height - Measuring Unit'      => [$isRequired, 'Height - Measuring Unit'],
        'Width'                        => [$isRequired, 'Width'],
        'Width - Measuring Unit'       => [$isRequired, 'Width - Measuring Unit'],
        'Main Image URL'               => [$isRequired, 'Main Image URL']
    ];

    $invalidRows = [];
    $flipkartOutputStartRow = 5;
    $validRowCounter = $flipkartOutputStartRow;

    // --- Step 4: Process each CSV row ---
    foreach ($inputData as $row) {
        $errorList = [];
        $rowValues = [];

        foreach ($mapping as $targetHeader => $mapDetails) {
            list($isRequiredField, $baseFileHeader) = $mapDetails;
            $value = isset($row[$baseFileHeader]) ? $row[$baseFileHeader] : null;
            $rowValues[$baseFileHeader] = $value;

            if ($isRequiredField && (is_null($value) || trim((string)$value) === '' || strtolower((string)$value) === 'nan')) {
                $errorList[] = "Missing required value in column '$baseFileHeader'";
            }

            if (
                $baseFileHeader === 'MRP (INR)' || $baseFileHeader === 'Your selling price (INR)' || $baseFileHeader === 'Procurement SLA (DAY)' ||
                $baseFileHeader === 'Stock' || $baseFileHeader === 'Local delivery charge (INR)' || $baseFileHeader === 'Zonal delivery charge (INR)' ||
                $baseFileHeader === 'National delivery charge (INR)'
            ) {
                if (!IsPositiveInteger($value)) {
                    $errorList[] = "Column '$baseFileHeader' must be a positive integer; got '$value'";
                }
            }

            if ($baseFileHeader === 'Fullfilment by') {
                $trimmedVal = trim((string)$value);
                if (strtolower($trimmedVal) === strtolower(FULLFILMENT_BY_VALUE)) {
                    $value = FULLFILMENT_BY_VALUE;
                } else {
                    $errorList[] = "Column 'Fullfilment by' must be '" . FULLFILMENT_BY_VALUE . "'; got '$value'";
                }
            }

            if ($baseFileHeader === 'Shipping provider') {
                $trimmedVal = trim((string)$value);
                if ($trimmedVal !== SHIPPING_PROVIDER_VALUE) {
                    $errorList[] = "Column 'Shipping provider' must be '" . SHIPPING_PROVIDER_VALUE . "'; got '$value'";
                }
            }

            if (
                $baseFileHeader === 'Height (CM)' || $baseFileHeader === 'Weight (KG)' || $baseFileHeader === 'Breadth (CM)' ||
                $baseFileHeader === 'Length (CM)' || $baseFileHeader === 'Height' || $baseFileHeader === 'Width'
            ) {
                if (!is_numeric($value)) {
                    $errorList[] = "Column '$baseFileHeader' must be a number (int or decimal); got '$value'";
                }
            }

            if ($baseFileHeader === 'Country Of Origin') {
                $normalized = normalizeCountry($value);
                if ($normalized === false) {
                    $errorList[] = "Column 'Country Of Origin' must be a valid country name; got '$value'";
                } else {
                    $value = $normalized;
                }
            }

            if ($baseFileHeader === 'Tax Code') {
                if (!in_array(trim((string)$value), $allowedTaxCodes, true)) {
                    $errorList[] = "Column 'Tax Code' must be one of " . json_encode($allowedTaxCodes) . "; got '$value'";
                }
            }

            if ($baseFileHeader === 'Brand') {
                if (!in_array(trim((string)$value), $allowedBrands, true)) {
                    $errorList[] = "Column 'Brand' must be one of " . json_encode($allowedBrands) . "; got '$value'";
                }
            }

            if ($baseFileHeader === 'Model Name') {
                // Additional validation if needed.
            }

            if ($baseFileHeader === 'Brand Color') {
                // Additional validation if needed.
            }

            if ($baseFileHeader === 'Color') {
                // Additional validation if needed.
            }

            if ($baseFileHeader === 'Style Code') {
                // Additional validation if needed.
            }

            if ($baseFileHeader === 'Type') {
                if (!in_array(trim((string)$value), $allowedTypes, true)) {
                    $errorList[] = "Column 'Type' must be one of " . json_encode($allowedTypes) . "; got '$value'";
                }
            }

            if ($baseFileHeader === 'Ideal For') {
                if (!in_array(trim((string)$value), $allowedIdealFor, true)) {
                    $errorList[] = "Column 'Ideal For' must be one of " . json_encode($allowedIdealFor) . "; got '$value'";
                }
            }

            if ($baseFileHeader === 'Occasion') {
                if (!in_array(trim((string)$value), $allowedOccasions, true)) {
                    $errorList[] = "Column 'Occasion' must be one of " . json_encode($allowedOccasions) . "; got '$value'";
                }
            }

            if ($baseFileHeader === 'Material') {
                if (!in_array(trim((string)$value), $allowedMaterials, true)) {
                    $errorList[] = "Column 'Material' must be one of " . json_encode($allowedMaterials) . "; got '$value'";
                }
            }

            if ($baseFileHeader === 'Main Image URL') {
                if (!IsValidUrl($value)) {
                    $errorList[] = "Column 'Main Image URL' must be a valid URL; got '$value'";
                }
            }
            // Save the updated value back.
            $row[$baseFileHeader] = $value;
        }

        // Cross-field validation: for example, ensure 'Brand' and 'Model Name' are not identical.
        $brand = isset($row['Brand']) ? $row['Brand'] : null;
        $model = isset($row['Model Name']) ? $row['Model Name'] : null;
        if (!is_null($brand) && !is_null($model) && trim((string)$brand) === trim((string)$model)) {
            $errorList[] = "Columns 'Brand' and 'Model Name' cannot have the same value";
        }

        if (empty($errorList)) {
            foreach ($mapping as $targetHeader => $mapDetails) {
                list(, $baseFileHeader) = $mapDetails;
                $value = isset($row[$baseFileHeader]) ? $row[$baseFileHeader] : null;
                // Since template headers match the CSV headers, look up the target column letter.
                if (isset($templateHeaderMap[$targetHeader])) {
                    $targetColLetter = $templateHeaderMap[$targetHeader];
                    $cellCoordinate = $targetColLetter . $validRowCounter;
                    $flipkartSheet->setCellValue($cellCoordinate, $value);
                } else {
                    $errorList[] = "Template header '$targetHeader' not found.";
                }
            }
            $validRowCounter++;
        } else {
            $rowRecord = [];
            foreach ($row as $baseFileHeader => $cellValue) {
                $rowRecord[$baseFileHeader] = $cellValue;
            }
            $rowRecord['Validation Errors'] = implode(', ', $errorList);
            $invalidRows[] = $rowRecord;
        }
    }

    $flipkartOutputPath = 'C_sling-bag_filled.xlsx';
    try {
        $writer = IOFactory::createWriter($flipkartTemplateSpreadsheet, 'Xlsx');
        $writer->save($flipkartOutputPath);
        echo "✅ Valid rows have been filled into the template starting from row $flipkartOutputStartRow in the sling_bag tab.\n";
    } catch (Exception $e) {
        error_log("Error saving filled workbook: " . $e->getMessage());
        echo "Error: Unable to save filled workbook.\n";
    }

    if (!empty($invalidRows)) {
        try {
            $invalidSpreadsheet = new Spreadsheet();
            $invalidSheet = $invalidSpreadsheet->getActiveSheet();

            $reportHeaders = array_keys($invalidRows[0]);
            $colIndex = 1;
            foreach ($reportHeaders as $header) {
                $cellCoordinate = Coordinate::stringFromColumnIndex($colIndex) . '1';
                $invalidSheet->setCellValue($cellCoordinate, $header);
                $colIndex++;
            }

            $rowIndex = 2;
            foreach ($invalidRows as $rowData) {
                $colIndex = 1;
                foreach ($reportHeaders as $header) {
                    $cellCoordinate = Coordinate::stringFromColumnIndex($colIndex) . $rowIndex;
                    $invalidSheet->setCellValue($cellCoordinate, $rowData[$header]);
                    $colIndex++;
                }
                $rowIndex++;
            }

            $invalidReportPath = 'invalid_data_report.xlsx';
            $invalidWriter = IOFactory::createWriter($invalidSpreadsheet, 'Xlsx');
            $invalidWriter->save($invalidReportPath);
            echo "⚠️ Invalid data report generated: " . $invalidReportPath . "\n";

            $invalidSpreadsheet->disconnectWorksheets();
            unset($invalidSpreadsheet);
        } catch (Exception $e) {
            error_log("Error generating invalid data report: " . $e->getMessage());
            echo "Error: Unable to generate invalid data report.\n";
        }
    } else {
        echo "✅ All rows passed validation. No invalid data report generated.\n";
    }

    $baseInputSpreadsheet->disconnectWorksheets();
    unset($baseInputSpreadsheet);

    $flipkartTemplateSpreadsheet->disconnectWorksheets();
    unset($flipkartTemplateSpreadsheet);
}

if (__FILE__ === realpath($_SERVER['SCRIPT_FILENAME'])) {
    main();
}
