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
    $allowed_AD = ['GST_0', 'GST_12', 'GST_18', 'GST_3', 'GST_5', 'GST_APPAREL'];
    $allowed_AI = [
        'Beige', 'Black', 'Blue', 'Brown', 'Clear', 'Gold', 'Green', 'Grey',
        'Khaki', 'Maroon', 'Multicolor', 'Orange', 'Pink', 'Purple', 'Red',
        'Silver', 'Tan', 'White', 'Yellow'
    ];
    $allowed_AK = [
        'Clutch', 'Hand-held Bag', 'Hobo', 'Messenger Bag', 'Satchel',
        'Shoulder Bag', 'Sling Bag', 'Tote'
    ];
    $allowed_AL = ['Boys', 'Boys & Girls', 'Girls', 'Men', 'Men & Women', 'Women'];
    $allowed_AM = ['Casual', 'Evening/Party', 'Formal', 'Sports'];
    $allowed_AN = [
        'Acrylic', 'Beads', 'Brocade', 'Canvas', 'Cotton', 'Denim', 'Fabric',
        'Flex', 'Genuine Leather', 'Juco', 'Jute', 'Leatherette', 'Metal',
        'Natural Fibre', 'PU', 'Plastic', 'Polyester', 'Rexine', 'Satin',
        'Silicon', 'Silk', 'Synthetic Leather', 'Tyvek', 'Velvet', 'Wood', 'Wool'
    ];

    // Define variables for required flag.
    $isRequired = true;
    $isNotRequired = false;

    // --- Step 1: Read the CSV data (using header names) ---
    $csv_file_path = 'sample_data.csv'; // Replace with your CSV file path

    try {
        $reader = IOFactory::createReader('Csv');
        // Optionally adjust CSV settings (delimiter, enclosure, etc.)
        $reader->setReadDataOnly(true);
        $spreadsheet = $reader->load($csv_file_path);
    } catch (Exception $e) {
        error_log("Error loading CSV file: " . $e->getMessage());
        echo "Error: Unable to load CSV file.\n";
        return;
    }

    $worksheet = $spreadsheet->getActiveSheet();
    $rows = $worksheet->toArray(null, true, true, true);

    // Assume the first row contains headers.
    $headers = array_map('trim', array_values(reset($rows)));
    // Remove the header row from data.
    array_shift($rows);

    // Convert each row to an associative array using header names.
    $data = [];
    foreach ($rows as $row) {
        $assoc = [];
        $colIndex = 0;
        foreach ($headers as $header) {
            // Use column letters from PhpSpreadsheet for any fallback if needed.
            $letter = Coordinate::stringFromColumnIndex($colIndex + 1);
            $assoc[$header] = isset($row[$letter]) ? $row[$letter] : null;
            $colIndex++;
        }
        $data[] = $assoc;
    }

    // --- Step 2: Open the Excel template ---
    $template_path = 'C_sling-bag_fd927b15e6244645_1703-2438FK_REQH2ILIQXHAH.xlsx';

    try {
        $templateSpreadsheet = IOFactory::load($template_path);
    } catch (Exception $e) {
        error_log("Error loading template file: " . $e->getMessage());
        echo "Error: Unable to load template file.\n";
        return;
    }

    $sheet = $templateSpreadsheet->getSheetByName('sling_bag');

    // --- Step 3: Define the mapping ---
    // Mapping: target template column letter => [is_required, source CSV header title]
    $mapping = [
        'G'  => [$isRequired, 'Seller SKU ID'],
        'J'  => [$isRequired, 'MRP (INR)'],
        'K'  => [$isRequired, 'Your selling price (INR)'],
        'L'  => [$isRequired, 'Fullfilment by'],
        'N'  => [$isRequired, 'Procurement SLA (DAY)'],
        'O'  => [$isRequired, 'Stock'],
        'P'  => [$isRequired, 'Shipping provider'],
        'Q'  => [$isRequired, 'Local delivery charge (INR)'],
        'R'  => [$isRequired, 'Zonal delivery charge (INR)'],
        'S'  => [$isRequired, 'National delivery charge (INR)'],
        'T'  => [$isRequired, 'Height (CM)'],
        'U'  => [$isRequired, 'Weight (KG)'],
        'V'  => [$isRequired, 'Breadth (CM)'],
        'W'  => [$isRequired, 'Length (CM)'],
        'Z'  => [$isRequired, 'Country Of Origin'],
        'AA' => [$isRequired, 'Manufacturer Details'],
        'AB' => [$isRequired, 'Packer Details'],
        'AD' => [$isRequired, 'Tax Code'],
        'AF' => [$isRequired, 'Brand'],
        'AG' => [$isRequired, 'Model Name'],
        'AH' => [$isRequired, 'Brand Color'],
        'AI' => [$isRequired, 'Color'],
        'AJ' => [$isRequired, 'Style Code'],
        'AK' => [$isRequired, 'Type'],
        'AL' => [$isRequired, 'Ideal For'],
        'AM' => [$isRequired, 'Occasion'],
        'AN' => [$isRequired, 'Material'],
        'AO' => [$isRequired, 'Pack of'],
        'AP' => [$isRequired, 'Height'],
        'AQ' => [$isRequired, 'Height - Measuring Unit'],
        'AR' => [$isRequired, 'Width'],
        'AS' => [$isRequired, 'Width - Measuring Unit'],
        'AT' => [$isRequired, 'Main Image URL']
    ];

    $invalid_data_rows = [];
    $start_row = 5;
    $valid_row_counter = $start_row;

    // --- Step 4: Process each row from the CSV ---
    foreach ($data as $row) {
        $error_list = [];
        $row_values = [];

        foreach ($mapping as $target_col => $mapDetails) {
            list($is_required, $header) = $mapDetails;
            $value = isset($row[$header]) ? $row[$header] : null;
            $row_values[$header] = $value;

            // Check required fields.
            if ($is_required && (is_null($value) || trim((string)$value) === '' || strtolower((string)$value) === 'nan')) {
                $error_list[] = "Missing required value in column '$header'";
            }

            // Column-specific validations.
            if (
                $header === 'MRP (INR)' || $header === 'Your selling price (INR)' || $header === 'Procurement SLA (DAY)' ||
                $header === 'Stock' || $header === 'Local delivery charge (INR)' || $header === 'Zonal delivery charge (INR)' ||
                $header === 'National delivery charge (INR)'
            ) {
                if (!IsPositiveInteger($value)) {
                    $error_list[] = "Column '$header' must be a positive integer; got '$value'";
                }
            }

            if ($header === 'Fullfilment by') {
                $trimmedValue = trim((string)$value);
                // If value is "seller" in lowercase, convert it to "Seller"
                if ($trimmedValue === 'seller') {
                    $value = 'Seller';
                } elseif ($trimmedValue !== 'Seller') {
                    $error_list[] = "Column 'Fullfilment by' must be 'Seller'; got '$value'";
                }
            }

            if ($header === 'Flipkart') {
                // Not applicable here; see below.
            }

            if (
                $header === 'Height (CM)' || $header === 'Weight (KG)' || $header === 'Breadth (CM)' ||
                $header === 'Length (CM)' || $header === 'Height' || $header === 'Width'
            ) {
                if (!is_numeric($value)) {
                    $error_list[] = "Column '$header' must be a number (int or decimal); got '$value'";
                }
            }

            if ($header === 'Country Of Origin') {
                $normalized = normalizeCountry($value);
                if ($normalized === false) {
                    $error_list[] = "Column 'Country Of Origin' must be a valid country name; got '$value'";
                } else {
                    $value = $normalized;
                }
            }

            if ($header === 'Tax Code') {
                if (!in_array(trim((string)$value), $allowed_AD, true)) {
                    $error_list[] = "Column 'Tax Code' must be one of " . json_encode($allowed_AD) . "; got '$value'";
                }
            }

            if ($header === 'Brand') {
                if (!in_array(trim((string)$value), $allowed_AI, true)) {
                    $error_list[] = "Column 'Brand' must be one of " . json_encode($allowed_AI) . "; got '$value'";
                }
            }

            if ($header === 'Model Name') {
                // Additional validation can be added here if needed.
            }

            if ($header === 'Brand Color') {
                // Additional validation if needed.
            }

            if ($header === 'Color') {
                // Additional validation if needed.
            }

            if ($header === 'Style Code') {
                // Additional validation if needed.
            }

            if ($header === 'Type') {
                if (!in_array(trim((string)$value), $allowed_AK, true)) {
                    $error_list[] = "Column 'Type' must be one of " . json_encode($allowed_AK) . "; got '$value'";
                }
            }

            if ($header === 'Ideal For') {
                if (!in_array(trim((string)$value), $allowed_AL, true)) {
                    $error_list[] = "Column 'Ideal For' must be one of " . json_encode($allowed_AL) . "; got '$value'";
                }
            }

            if ($header === 'Occasion') {
                if (!in_array(trim((string)$value), $allowed_AM, true)) {
                    $error_list[] = "Column 'Occasion' must be one of " . json_encode($allowed_AM) . "; got '$value'";
                }
            }

            if ($header === 'Material') {
                if (!in_array(trim((string)$value), $allowed_AN, true)) {
                    $error_list[] = "Column 'Material' must be one of " . json_encode($allowed_AN) . "; got '$value'";
                }
            }

            // For columns like "Shipping provider", "Manufacturer Details", etc.,
            // you can add additional validation as needed.

            // For "Main Image URL"
            if ($header === 'Main Image URL') {
                if (!IsValidUrl($value)) {
                    $error_list[] = "Column 'Main Image URL' must be a valid URL; got '$value'";
                }
            }
            // Save the possibly updated value back to the row's associative array.
            $row[$header] = $value;
        }

        // Cross-field validation for columns "Brand" (AF) and "Model Name" (AG), for example.
        // Adjust if needed (currently not specified in mapping).
        // For example:
        $brand = isset($row['Brand']) ? $row['Brand'] : null;
        $model = isset($row['Model Name']) ? $row['Model Name'] : null;
        if (!is_null($brand) && !is_null($model) && trim((string)$brand) === trim((string)$model)) {
            $error_list[] = "Columns 'Brand' and 'Model Name' cannot have the same value";
        }

        if (empty($error_list)) {
            foreach ($mapping as $target_col => $mapDetails) {
                list(, $header) = $mapDetails;
                $value = isset($row[$header]) ? $row[$header] : null;
                $cellCoordinate = $target_col . $valid_row_counter;
                $sheet->setCellValue($cellCoordinate, $value);
            }
            $valid_row_counter++;
        } else {
            $row_dict = [];
            foreach ($row as $header => $cellValue) {
                $row_dict[$header] = $cellValue;
            }
            $row_dict['Validation Errors'] = implode(', ', $error_list);
            $invalid_data_rows[] = $row_dict;
        }
    }

    $output_path = 'C_sling-bag_filled.xlsx';
    try {
        $writer = IOFactory::createWriter($templateSpreadsheet, 'Xlsx');
        $writer->save($output_path);
        echo "✅ Valid rows have been filled into the template starting from row 5 in the sling_bag tab.\n";
    } catch (Exception $e) {
        error_log("Error saving filled workbook: " . $e->getMessage());
        echo "Error: Unable to save filled workbook.\n";
    }

    if (!empty($invalid_data_rows)) {
        try {
            $invalidSpreadsheet = new Spreadsheet();
            $invalidSheet = $invalidSpreadsheet->getActiveSheet();

            $headers = array_keys($invalid_data_rows[0]);
            $colIndex = 1;
            foreach ($headers as $header) {
                $cellCoordinate = Coordinate::stringFromColumnIndex($colIndex) . '1';
                $invalidSheet->setCellValue($cellCoordinate, $header);
                $colIndex++;
            }

            $rowIndex = 2;
            foreach ($invalid_data_rows as $rowData) {
                $colIndex = 1;
                foreach ($headers as $header) {
                    $cellCoordinate = Coordinate::stringFromColumnIndex($colIndex) . $rowIndex;
                    $invalidSheet->setCellValue($cellCoordinate, $rowData[$header]);
                    $colIndex++;
                }
                $rowIndex++;
            }

            $invalid_report_path = 'invalid_data_report.xlsx';
            $invalidWriter = IOFactory::createWriter($invalidSpreadsheet, 'Xlsx');
            $invalidWriter->save($invalid_report_path);
            echo "⚠️ Invalid data report generated: " . $invalid_report_path . "\n";

            $invalidSpreadsheet->disconnectWorksheets();
            unset($invalidSpreadsheet);
        } catch (Exception $e) {
            error_log("Error generating invalid data report: " . $e->getMessage());
            echo "Error: Unable to generate invalid data report.\n";
        }
    } else {
        echo "✅ All rows passed validation. No invalid data report generated.\n";
    }

    $spreadsheet->disconnectWorksheets();
    unset($spreadsheet);

    $templateSpreadsheet->disconnectWorksheets();
    unset($templateSpreadsheet);
}

if (__FILE__ === realpath($_SERVER['SCRIPT_FILENAME'])) {
    main();
}
