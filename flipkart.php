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
 *
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
 * Checks if a given value is numeric.
 *
 * @param mixed $value
 *
 * @return bool
 */
function IsNumber($value): bool
{
    return is_numeric($value);
}

/**
 * Checks if a value is a valid number (integer or decimal).
 *
 * @param mixed $value
 *
 * @return bool
 */
function IsValidDecimalOrInt($value): bool
{
    return IsNumber($value);
}

/**
 * Checks if a value is one of the allowed length units.
 *
 * @param mixed $value
 *
 * @return bool
 */
function IsValidLengthUnit($value): bool
{
    $allowedUnits = ['cm', 'mm', 'inch'];
    return in_array((string)$value, $allowedUnits, true);
}

/**
 * Checks if a given country name exists in the predefined list.
 *
 * @param mixed $value
 *
 * @return bool
 */
function IsValidCountry($value): bool
{
    $country = trim((string)$value);
    return in_array($country, COUNTRIES, true);
}

/**
 * Checks if a value is a valid URL.
 *
 * @param mixed $value
 *
 * @return bool
 */
function IsValidUrl($value): bool
{
    $value = trim((string)$value);
    return filter_var($value, FILTER_VALIDATE_URL) !== false;
}

/**
 * Main execution function.
 *
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
    $isRequired    = true;
    $isNotRequired = false;

    // --- Step 1: Read the XLSX data (ignoring header names) ---
    $xlsx_file_path = 'sample_data (1).xlsx'; // Replace with your XLSX file path

    try {
        $reader      = IOFactory::createReader('Xlsx');
        $reader->setReadDataOnly(true);
        $spreadsheet = $reader->load($xlsx_file_path);
    } catch (Exception $e) {
        error_log("Error loading XLSX file: " . $e->getMessage());
        echo "Error: Unable to load XLSX file.\n";
        return;
    }

    $worksheet = $spreadsheet->getActiveSheet();
    // PhpSpreadsheet reads data into an array with keys as column letters (A, B, C, …)
    $data = $worksheet->toArray(null, true, true, true);

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
    // Mapping: target template column letter => [is_required, source Excel column letter]
    $mapping = [
        'G'  => [$isRequired, 'G'],
        'J'  => [$isRequired, 'J'],  // Must be positive integer
        'K'  => [$isRequired, 'K'],  // Must be positive integer
        'L'  => [$isRequired, 'L'],  // Must be exactly "Seller"
        'N'  => [$isRequired, 'N'],  // Must be positive integer
        'O'  => [$isRequired, 'O'],  // Must be positive integer
        'P'  => [$isRequired, 'P'],  // Must be exactly "Flipkart"
        'Q'  => [$isRequired, 'Q'],  // Must be positive integer
        'R'  => [$isRequired, 'R'],  // Must be positive integer
        'S'  => [$isRequired, 'S'],  // Must be positive integer
        'T'  => [$isRequired, 'T'],  // Can be int or decimal
        'U'  => [$isRequired, 'U'],  // Can be int or decimal
        'V'  => [$isRequired, 'V'],  // Can be int or decimal
        'W'  => [$isRequired, 'W'],  // Can be int or decimal
        'Z'  => [$isRequired, 'Z'],  // Valid country (must be in the list)
        'AA' => [$isRequired, 'AA'],
        'AB' => [$isRequired, 'AB'],
        'AD' => [$isRequired, 'AD'], // Only allowed GST values
        'AF' => [$isRequired, 'AF'],
        'AG' => [$isRequired, 'AG'], // AF and AG must not be the same
        'AH' => [$isRequired, 'AH'],
        'AI' => [$isRequired, 'AI'], // Must be one of allowed_AI
        'AJ' => [$isRequired, 'AJ'],
        'AK' => [$isRequired, 'AK'], // Must be one of allowed_AK
        'AL' => [$isRequired, 'AL'], // Must be one of allowed_AL
        'AM' => [$isRequired, 'AM'], // Must be one of allowed_AM
        'AN' => [$isRequired, 'AN'], // Must be one of allowed_AN
        'AO' => [$isRequired, 'AO'], // Must be a number
        'AP' => [$isRequired, 'AP'], // Can be int or decimal
        'AQ' => [$isRequired, 'AQ'], // Must be one of allowed units ("cm", "mm", "inch")
        'AR' => [$isRequired, 'AR'], // Can be int or decimal
        'AS' => [$isRequired, 'AS'], // Must be one of allowed units ("cm", "mm", "inch")
        'AT' => [$isRequired, 'AT']  // Must be a valid URL
    ];

    // --- Step 4: Prepare for invalid data tracking and valid row counter ---
    $invalid_data_rows = []; // Rows that fail validation
    $start_row         = 5;  // Row in the template where data will be written
    $valid_row_counter = $start_row; // Counter for valid rows in the template

    // --- Step 5: Process each row from the input file ---
    // $data is an array of rows. Each row is an associative array with keys as column letters.
    foreach ($data as $row) {
        $error_list = [];
        $row_values = [];

        // Validate each mapped field.
        foreach ($mapping as $target_col => $mapDetails) {
            list($is_required, $source_letter) = $mapDetails;
            $value                      = isset($row[$source_letter]) ? $row[$source_letter] : null;
            $row_values[$source_letter] = $value;

            // Check required fields.
            if ($is_required && (is_null($value) || trim((string)$value) === '' || strtolower((string)$value) === 'nan')) {
                $error_list[] = "Missing required value in column $source_letter";
            }

            // Column-specific validations.
            if ($source_letter === 'J') {
                if (!IsPositiveInteger($value)) {
                    $error_list[] = "Column J must be a positive integer; got '$value'";
                }
            }

            if ($source_letter === 'K') {
                if (!IsPositiveInteger($value)) {
                    $error_list[] = "Column K must be a positive integer; got '$value'";
                }
            }

            if ($source_letter === 'L') {
                $trimmedValue = trim((string)$value);
                // If value is "seller" in lowercase, convert it to "Seller"
                if ($trimmedValue === 'seller') {
                    $value = 'Seller';
                } elseif ($trimmedValue !== 'Seller') {
                    $error_list[] = "Column L must be 'Seller'; got '$value'";
                }
            }

            if ($source_letter === 'N') {
                if (!IsPositiveInteger($value)) {
                    $error_list[] = "Column N must be a positive integer; got '$value'";
                }
            }

            if ($source_letter === 'O') {
                if (!IsPositiveInteger($value)) {
                    $error_list[] = "Column O must be a positive integer; got '$value'";
                }
            }

            if ($source_letter === 'P') {
                if (trim((string)$value) !== 'Flipkart') {
                    $error_list[] = "Column P must be 'Flipkart'; got '$value'";
                }
            }

            if (in_array($source_letter, ['Q', 'R', 'S'], true)) {
                if (!IsPositiveInteger($value)) {
                    $error_list[] = "Column $source_letter must be a positive integer; got '$value'";
                }
            }

            if (in_array($source_letter, ['T', 'U', 'V', 'W'], true)) {
                if (!IsValidDecimalOrInt($value)) {
                    $error_list[] = "Column $source_letter must be a number (int or decimal); got '$value'";
                }
            }

            if ($source_letter === 'Z') {
                if (!IsValidCountry($value)) {
                    $error_list[] = "Column Z must be a valid country name (must be in the predefined list); got '$value'";
                }
            }

            if ($source_letter === 'AD') {
                if (!in_array(trim((string)$value), $allowed_AD, true)) {
                    $error_list[] = "Column AD must be one of " . json_encode($allowed_AD) .
                        "; got '$value'";
                }
            }

            if ($source_letter === 'AI') {
                if (!in_array(trim((string)$value), $allowed_AI, true)) {
                    $error_list[] = "Column AI must be one of " . json_encode($allowed_AI) .
                        "; got '$value'";
                }
            }

            if ($source_letter === 'AK') {
                if (!in_array(trim((string)$value), $allowed_AK, true)) {
                    $error_list[] = "Column AK must be one of " . json_encode($allowed_AK) .
                        "; got '$value'";
                }
            }

            if ($source_letter === 'AL') {
                if (!in_array(trim((string)$value), $allowed_AL, true)) {
                    $error_list[] = "Column AL must be one of " . json_encode($allowed_AL) .
                        "; got '$value'";
                }
            }

            if ($source_letter === 'AM') {
                if (!in_array(trim((string)$value), $allowed_AM, true)) {
                    $error_list[] = "Column AM must be one of " . json_encode($allowed_AM) .
                        "; got '$value'";
                }
            }

            if ($source_letter === 'AN') {
                if (!in_array(trim((string)$value), $allowed_AN, true)) {
                    $error_list[] = "Column AN must be one of " . json_encode($allowed_AN) .
                        "; got '$value'";
                }
            }

            if ($source_letter === 'AO') {
                if (!IsNumber($value)) {
                    $error_list[] = "Column AO must be a number; got '$value'";
                }
            }

            if (in_array($source_letter, ['AP', 'AR'], true)) {
                if (!IsValidDecimalOrInt($value)) {
                    $error_list[] = "Column $source_letter must be a number (int or decimal); got '$value'";
                }
            }

            if (in_array($source_letter, ['AQ', 'AS'], true)) {
                if (!IsValidLengthUnit($value)) {
                    $error_list[] = "Column $source_letter must be one of 'cm', 'mm', 'inch'; got '$value'";
                }
            }

            if ($source_letter === 'AT') {
                if (!IsValidUrl($value)) {
                    $error_list[] = "Column AT must be a valid URL; got '$value'";
                }
            }
        }

        // Cross-field validation: AF and AG must not have the same value.
        $af_val = isset($row_values['AF']) ? $row_values['AF'] : null;
        $ag_val = isset($row_values['AG']) ? $row_values['AG'] : null;
        if (!is_null($af_val) && !is_null($ag_val)) {
            if (trim((string)$af_val) === trim((string)$ag_val)) {
                $error_list[] = 'Columns AF and AG cannot have the same value';
            }
        }

        // If row is valid, write it to the template; otherwise, add it to invalid data report.
        if (empty($error_list)) {
            // Write each mapped field into the template in the row indicated by $valid_row_counter.
            foreach ($mapping as $target_col => $mapDetails) {
                list(, $source_letter) = $mapDetails;
                $value          = isset($row[$source_letter]) ? $row[$source_letter] : null;
                $cellCoordinate = $target_col . $valid_row_counter;
                $sheet->setCellValue($cellCoordinate, $value);
            }
            $valid_row_counter++;
        } else {
            // Create an array for the row using Excel letters for columns.
            $row_dict = [];
            foreach ($row as $colLetter => $cellValue) {
                $row_dict[$colLetter] = $cellValue;
            }
            $row_dict['Validation Errors'] = implode(', ', $error_list);
            $invalid_data_rows[] = $row_dict;
        }
    }

    // --- Step 6: Save the updated workbook with only valid rows ---
    $output_path = 'C_sling-bag_filled.xlsx';
    try {
        $writer = IOFactory::createWriter($templateSpreadsheet, 'Xlsx');
        $writer->save($output_path);
        echo "✅ Valid rows have been filled into the template starting from row 5 in the sling_bag tab.\n";
    } catch (Exception $e) {
        error_log("Error saving filled workbook: " . $e->getMessage());
        echo "Error: Unable to save filled workbook.\n";
    }

    // --- Step 7: Create a report for invalid rows (if any) ---
    if (!empty($invalid_data_rows)) {
        try {
            // Create a new spreadsheet for the invalid report.
            $invalidSpreadsheet = new Spreadsheet();
            $invalidSheet       = $invalidSpreadsheet->getActiveSheet();

            // Write header row.
            $headers  = array_keys($invalid_data_rows[0]);
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
            $invalidWriter      = IOFactory::createWriter($invalidSpreadsheet, 'Xlsx');
            $invalidWriter->save($invalid_report_path);
            echo "⚠️ Invalid data report generated: " . $invalid_report_path . "\n";

            // Free memory for invalid report.
            $invalidSpreadsheet->disconnectWorksheets();
            unset($invalidSpreadsheet);
        } catch (Exception $e) {
            error_log("Error generating invalid data report: " . $e->getMessage());
            echo "Error: Unable to generate invalid data report.\n";
        }
    } else {
        echo "✅ All rows passed validation. No invalid data report generated.\n";
    }

    // Free memory for main spreadsheets.
    $spreadsheet->disconnectWorksheets();
    unset($spreadsheet);

    $templateSpreadsheet->disconnectWorksheets();
    unset($templateSpreadsheet);
}

// Only call main() if this file is executed directly.
if (__FILE__ === realpath($_SERVER['SCRIPT_FILENAME'])) {
    main();
}
