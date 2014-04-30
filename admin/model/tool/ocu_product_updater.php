<?php

/**
 * OpenCart Ukrainian Community
 *
 * LICENSE
 *
 * This source file is subject to the GNU General Public License, Version 3
 * It is also available through the world-wide-web at this URL:
 * http://www.gnu.org/copyleft/gpl.html
 * If you did not receive a copy of the license and are unable to
 * obtain it through the world-wide-web, please send an email

 *
 * This product based on export/import module by maxzon.ru
 * Contains other components. That are distributed on other licenses, specified in theirs source code.
 *
 * @category   OpenCart
 * @package    OCU Product Updater
 * @copyright  Copyright (c) 2011 Eugene Lifescale by OpenCart Ukrainian Community (http://opencart-ukraine.tumblr.com)
 * @license    http://www.gnu.org/copyleft/gpl.html     GNU General Public License, Version 3
 */


static $config = NULL;
static $log = NULL;

// Error Handler
function error_handler_for_export($errno, $errstr, $errfile, $errline)
{
    global $config;
    global $log;

    switch ($errno) {
        case E_NOTICE:
        case E_USER_NOTICE:
            $errors = "Notice";
            break;
        case E_WARNING:
        case E_USER_WARNING:
            $errors = "Warning";
            break;
        case E_ERROR:
        case E_USER_ERROR:
            $errors = "Fatal Error";
            break;
        default:
            $errors = "Unknown";
            break;
    }

    if (($errors=='Warning') || ($errors=='Unknown')) {
        return true;
    }

    if ($config->get('config_error_display')) {
        echo '<b>' . $errors . '</b>: ' . $errstr . ' in <b>' . $errfile . '</b> on line <b>' . $errline . '</b>';
    }

    if ($config->get('config_error_log')) {
        $log->write('PHP ' . $errors . ':  ' . $errstr . ' in ' . $errfile . ' on line ' . $errline);
    }

    return true;
}


function fatal_error_shutdown_handler_for_export()
{
    $last_error = error_get_last();
    if ($last_error['type'] === E_ERROR) {
        // fatal error
        error_handler_for_export(E_ERROR, $last_error['message'], $last_error['file'], $last_error['line']);
    }
}


class ModelToolOcuProductUpdater extends Model {


    function clean( &$str, $allowBlanks=FALSE )
    {
        $result = "";
        $n = strlen( $str );
        for ($m=0; $m<$n; $m++) {
            $ch = substr( $str, $m, 1 );
            if (($ch==" ") && (!$allowBlanks) || ($ch=="\n") || ($ch=="\r") || ($ch=="\t") || ($ch=="\0") || ($ch=="\x0B")) {
                continue;
            }
            $result .= $ch;
        }
        return $result;
    }


    function import( &$database, $sql )
    {
        foreach (explode(";\n", $sql) as $sql) {
            $sql = trim($sql);
            if ($sql) {
                $database->query($sql);
            }
        }
    }


    protected function getDefaultLanguageId( &$database )
    {
        $code = $this->config->get('config_language');
        $sql = "SELECT language_id FROM `".DB_PREFIX."language` WHERE code = '$code'";
        $result = $database->query( $sql );
        $languageId = 1;
        if ($result->rows) {
            foreach ($result->rows as $row) {
                $languageId = $row['language_id'];
                break;
            }
        }
        return $languageId;
    }


    protected function getDefaultWeightUnit()
    {
        $weightUnit = $this->config->get( 'config_weight_class' );
        return $weightUnit;
    }


    protected function getDefaultMeasurementUnit()
    {
        $measurementUnit = $this->config->get( 'config_length_class' );
        return $measurementUnit;
    }


    function storeProductsIntoDatabase( &$database, &$products )
    {
        // start transaction, remove products
        $database->query("START TRANSACTION;");

        // generate and execute SQL for storing the products
        foreach ($products as $product) {
            $price = $product['price'];
            $quantity = $product['quantity'];
            $sku = $database->escape($product['sku']);

            $database->query("UPDATE IGNORE `".DB_PREFIX."product` SET price = '$price', quantity = '$quantity' WHERE sku = '$sku' LIMIT 1");
        }

        // final update commit
        $database->query("COMMIT;");
        return TRUE;
    }


    protected function detect_encoding( $str )
    {
        // auto detect the character encoding of a string
        return mb_detect_encoding( $str, 'UTF-8,ISO-8859-15,ISO-8859-1,cp1251,KOI8-R' );
    }


    function uploadProducts( &$reader, &$database )
    {
        $data = $reader->getSheet(0);
        $products = array();
        $isFirstRow = TRUE;
        $k = $data->getHighestRow();
        for ($i=0; $i<$k; $i+=1) {
            if ($isFirstRow) {
                $isFirstRow = FALSE;
                continue;
            }
            $sku = trim($this->getCell($data,$i,1));
            $sku = htmlentities( $sku, ENT_QUOTES, $this->detect_encoding($sku) );
            $price = $this->getCell($data,$i,2);
            $quantity = $this->getCell($data,$i,3,'0');

            $product = array();
            $product['price'] = $price;
            $product['quantity'] = $quantity;
            $product['sku'] = $sku;
            $products[$sku] = $product;
        }
        return $this->storeProductsIntoDatabase( $database, $products );
    }


    function storeSpecialsIntoDatabase( &$database, &$specials )
       {
           $sql = "START TRANSACTION;\n";
           $sql .= "DELETE FROM `".DB_PREFIX."product_special`;\n";
           $this->import( $database, $sql );

           // find existing customer groups from the database
           $sql = "SELECT * FROM `".DB_PREFIX."customer_group`";
           $result = $database->query( $sql );
           $maxCustomerGroupId = 0;
           $customerGroups = array();
           foreach ($result->rows as $row) {
               $customerGroupId = $row['customer_group_id'];
               $name = $row['name'];
               if (!isset($customerGroups[$name])) {
                   $customerGroups[$name] = $customerGroupId;
               }
               if ($maxCustomerGroupId < $customerGroupId) {
                   $maxCustomerGroupId = $customerGroupId;
               }
           }

           // add additional customer groups into the database
           foreach ($specials as $special) {
               $name = $special['customer_group'];
               if (!isset($customerGroups[$name])) {
                   $maxCustomerGroupId += 1;
                   $sql  = "INSERT INTO `".DB_PREFIX."customer_group` (`customer_group_id`, `name`) VALUES ";
                   $sql .= "($maxCustomerGroupId, '$name')";
                   $sql .= ";\n";
                   $database->query($sql);
                   $customerGroups[$name] = $maxCustomerGroupId;
               }
           }

           // store product specials into the database
           $productSpecialId = 0;
           $first = TRUE;
           $sql = "INSERT INTO `".DB_PREFIX."product_special` (`product_special_id`,`product_id`,`customer_group_id`,`priority`,`price`,`date_start`,`date_end` ) VALUES ";
           foreach ($specials as $special) {
               $productSpecialId += 1;
               $productId = $special['product_id'];
               $name = $special['customer_group'];
               $customerGroupId = $customerGroups[$name];
               $priority = $special['priority'];
               $price = $special['price'];
               $dateStart = $special['date_start'];
               $dateEnd = $special['date_end'];
               $sql .= ($first) ? "\n" : ",\n";
               $first = FALSE;
               $sql .= "($productSpecialId,$productId,$customerGroupId,$priority,$price,'$dateStart','$dateEnd')";
           }
           if (!$first) {
               $database->query($sql);
           }

           $database->query("COMMIT;");
           return TRUE;
       }


       function uploadSpecials( &$reader, &$database )
       {
           $data = $reader->getSheet(1);
           $specials = array();
           $i = 0;
           $k = $data->getHighestRow();
           $isFirstRow = TRUE;
           for ($i=0; $i<$k; $i+=1) {
               if ($isFirstRow) {
                   $isFirstRow = FALSE;
                   continue;
               }
            $sku = trim($this->getCell($data,$i,1));
            $sku = htmlentities( $sku, ENT_QUOTES, $this->detect_encoding($sku) );
               if ($sku=="") {
                   continue;
               }
               $customerGroup = trim($this->getCell($data,$i,2));
               if ($customerGroup=="") {
                   continue;
               }

            $priority = $this->getCell($data,$i,3,'0');
            $price = $this->getCell($data,$i,4,'0');
            $dateStart = $this->getCell($data,$i,5,'0000-00-00');
            $dateEnd = $this->getCell($data,$i,6,'0000-00-00');

            $product = $database->query("SELECT product_id FROM `".DB_PREFIX."product` WHERE sku = '{$sku}' LIMIT 1");
            if ($product->num_rows) {
                $specials[$i] = array();
                $specials[$i]['product_id'] = $product->row['product_id'];
                $specials[$i]['customer_group'] = $customerGroup;
                $specials[$i]['priority'] = $priority;
                $specials[$i]['price'] = $price;
                $specials[$i]['date_start'] = $dateStart;
                $specials[$i]['date_end'] = $dateEnd;
            }
           }
           return $this->storeSpecialsIntoDatabase( $database, $specials );
       }


    function storeDiscountsIntoDatabase( &$database, &$discounts )
    {
       $sql = "START TRANSACTION;\n";
       $sql .= "DELETE FROM `".DB_PREFIX."product_discount`;\n";
       $this->import( $database, $sql );

       // find existing customer groups from the database
       $sql = "SELECT * FROM `".DB_PREFIX."customer_group`";
       $result = $database->query( $sql );
       $maxCustomerGroupId = 0;
       $customerGroups = array();
       foreach ($result->rows as $row) {
           $customerGroupId = $row['customer_group_id'];
           $name = $row['name'];
           if (!isset($customerGroups[$name])) {
               $customerGroups[$name] = $customerGroupId;
           }
           if ($maxCustomerGroupId < $customerGroupId) {
               $maxCustomerGroupId = $customerGroupId;
           }
       }

       // add additional customer groups into the database
       foreach ($discounts as $discount) {
           $name = $discount['customer_group'];
           if (!isset($customerGroups[$name])) {
               $maxCustomerGroupId += 1;
               $sql  = "INSERT INTO `".DB_PREFIX."customer_group` (`customer_group_id`, `name`) VALUES ";
               $sql .= "($maxCustomerGroupId, '$name')";
               $sql .= ";\n";
               $database->query($sql);
               $customerGroups[$name] = $maxCustomerGroupId;
           }
       }

       // store product discounts into the database
       $productDiscountId = 0;
       $first = TRUE;
       $sql = "INSERT INTO `".DB_PREFIX."product_discount` (`product_discount_id`,`product_id`,`customer_group_id`,`quantity`,`priority`,`price`,`date_start`,`date_end` ) VALUES ";
       foreach ($discounts as $discount) {
           $productDiscountId += 1;
           $productId = $discount['product_id'];
           $name = $discount['customer_group'];
           $customerGroupId = $customerGroups[$name];
           $quantity = $discount['quantity'];
           $priority = $discount['priority'];
           $price = $discount['price'];
           $dateStart = $discount['date_start'];
           $dateEnd = $discount['date_end'];
           $sql .= ($first) ? "\n" : ",\n";
           $first = FALSE;
           $sql .= "($productDiscountId,$productId,$customerGroupId,$quantity,$priority,$price,'$dateStart','$dateEnd')";
       }
       if (!$first) {
           $database->query($sql);
       }

       $database->query("COMMIT;");
       return TRUE;
    }


       function uploadDiscounts( &$reader, &$database )
       {
           $data = $reader->getSheet(2);
           $discounts = array();
           $i = 0;
           $k = $data->getHighestRow();
           $isFirstRow = TRUE;
           for ($i=0; $i<$k; $i+=1) {
               if ($isFirstRow) {
                   $isFirstRow = FALSE;
                   continue;
               }
            $sku = trim($this->getCell($data,$i,1));
            $sku = htmlentities( $sku, ENT_QUOTES, $this->detect_encoding($sku) );
               if ($sku=="") {
                   continue;
               }
               $customerGroup = trim($this->getCell($data,$i,2));
               if ($customerGroup=="") {
                   continue;
               }
               $quantity = $this->getCell($data,$i,3,'0');
               $priority = $this->getCell($data,$i,4,'0');
               $price = $this->getCell($data,$i,5,'0');
               $dateStart = $this->getCell($data,$i,6,'0000-00-00');
               $dateEnd = $this->getCell($data,$i,7,'0000-00-00');

            $product = $database->query("SELECT product_id FROM `".DB_PREFIX."product` WHERE sku = '{$sku}' LIMIT 1");
            if ($product->num_rows) {
                $discounts[$i] = array();
                $discounts[$i]['product_id'] = $product->row['product_id'];
                $discounts[$i]['customer_group'] = $customerGroup;
                $discounts[$i]['quantity'] = $quantity;
                $discounts[$i]['priority'] = $priority;
                $discounts[$i]['price'] = $price;
                $discounts[$i]['date_start'] = $dateStart;
                $discounts[$i]['date_end'] = $dateEnd;
            }
           }
           return $this->storeDiscountsIntoDatabase( $database, $discounts );
       }

    function getCell(&$worksheet,$row,$col,$default_val='')
    {
        $col -= 1; // we use 1-based, PHPExcel uses 0-based column index
        $row += 1; // we use 0-based, PHPExcel used 1-based row index
        return ($worksheet->cellExistsByColumnAndRow($col,$row)) ? $worksheet->getCellByColumnAndRow($col,$row)->getValue() : $default_val;
    }

    function validateHeading( &$data, &$expected )
    {
        $heading = array();
        $k = PHPExcel_Cell::columnIndexFromString( $data->getHighestColumn() );
        if ($k != count($expected)) {
            return FALSE;
        }
        $i = 0;
        for ($j=1; $j <= $k; $j+=1) {
            $heading[] = $this->getCell($data,$i,$j);
        }
        $valid = TRUE;
        for ($i=0; $i < count($expected); $i+=1) {
            if (!isset($heading[$i])) {
                $valid = FALSE;
                break;
            }
            if (strtolower($heading[$i]) != strtolower($expected[$i])) {
                $valid = FALSE;
                break;
            }
        }
        return $valid;
    }

    function validateProducts( &$reader )
    {
        $expectedProductHeading = array
        ( "sku", "price", "quantity");
        $data =& $reader->getSheet(0);
        return $this->validateHeading( $data, $expectedProductHeading );
    }

    function validateSpecials( &$reader )
       {
           $expectedSpecialsHeading = array
           ( "sku", "customer_group", "priority", "price", "date_start", "date_end" );
           $data =& $reader->getSheet(1);
           return $this->validateHeading( $data, $expectedSpecialsHeading );
       }


       function validateDiscounts( &$reader )
       {
           $expectedDiscountsHeading = array
           ( "sku", "customer_group", "quantity", "priority", "price", "date_start", "date_end" );
           $data =& $reader->getSheet(2);
           return $this->validateHeading( $data, $expectedDiscountsHeading );
       }

    function validateUpload( &$reader )
    {
        if ($reader->getSheetCount() != 3) {
            error_log(date('Y-m-d H:i:s - ', time()).$this->language->get( 'error_sheet_count' )."\n",3,DIR_LOGS."error.txt");
            return FALSE;
        }
        if (!$this->validateProducts( $reader )) {
            error_log(date('Y-m-d H:i:s - ', time()).$this->language->get('error_products_header')."\n",3,DIR_LOGS."error.txt");
            return FALSE;
        }
        if (!$this->validateSpecials( $reader )) {
            error_log(date('Y-m-d H:i:s - ', time()).$this->language->get('error_specials_header')."\n",3,DIR_LOGS."error.txt");
            return FALSE;
        }
        if (!$this->validateDiscounts( $reader )) {
            error_log(date('Y-m-d H:i:s - ', time()).$this->language->get('error_discounts_header')."\n",3,DIR_LOGS."error.txt");
            return FALSE;
        }
        return TRUE;
    }


    function clearCache()
    {
        $this->cache->delete('category');
        $this->cache->delete('category_description');
        $this->cache->delete('manufacturer');
        $this->cache->delete('product');
        $this->cache->delete('product_image');
        $this->cache->delete('product_option');
        $this->cache->delete('product_option_description');
        $this->cache->delete('product_option_value');
        $this->cache->delete('product_option_value_description');
        $this->cache->delete('product_to_category');
        $this->cache->delete('url_alias');
        $this->cache->delete('product_special');
        $this->cache->delete('product_discount');
    }


    function upload( $filename )
    {
        global $config;
        global $log;
        $config = $this->config;
        $log = $this->log;
        set_error_handler('error_handler_for_export',E_ALL);
        register_shutdown_function('fatal_error_shutdown_handler_for_export');
        $database =& $this->db;
        ini_set("memory_limit","512M");
        ini_set("max_execution_time",180);
        //set_time_limit( 60 );
        chdir( '../system/PHPExcel' );
        require_once( 'Classes/PHPExcel.php' );
        chdir( '../../admin' );
        $inputFileType = PHPExcel_IOFactory::identify($filename);
        $objReader = PHPExcel_IOFactory::createReader($inputFileType);
        $objReader->setReadDataOnly(true);
        $reader = $objReader->load($filename);
        $ok = $this->validateUpload( $reader );
        if (!$ok) {
            return FALSE;
        }
        $this->clearCache();

        $ok = $this->uploadProducts( $reader, $database );
        if (!$ok) {
            return FALSE;
        }
        $ok = $this->uploadSpecials( $reader, $database );
        if (!$ok) {
            return FALSE;
        }
        $ok = $this->uploadDiscounts( $reader, $database );
        if (!$ok) {
            return FALSE;
        }
        chdir( '../../..' );
        return $ok;
    }

    function populateProductsWorksheet( &$worksheet, &$database, &$boxFormat, &$textFormat )
    {
        // Set the column widths
        $j = 0;
        $worksheet->setColumn($j,$j++,max(strlen('sku'),10)+1);
        $worksheet->setColumn($j,$j++,max(strlen('price'),4)+1);
        $worksheet->setColumn($j,$j++,max(strlen('quantity'),4)+1);

        // The product headings row
        $i = 0;
        $j = 0;
        $worksheet->writeString( $i, $j++, 'sku', $boxFormat );
        $worksheet->writeString( $i, $j++, 'price', $boxFormat );
        $worksheet->writeString( $i, $j++, 'quantity', $boxFormat );
        $worksheet->setRow( $i, 30, $boxFormat );

        // The actual products data
        $i += 1;
        $j = 0;
        $query  = "SELECT ";
        $query .= "  p.product_id,";
        $query .= "  p.price,";
        $query .= "  p.sku,";
        $query .= "  p.quantity ";
        $query .= "FROM `".DB_PREFIX."product` p ";
        $query .= "LEFT JOIN `".DB_PREFIX."product_to_category` pc ON p.product_id=pc.product_id ";
        $query .= "GROUP BY p.product_id ";
        $query .= "ORDER BY p.product_id, pc.category_id; ";
        $result = $database->query( $query );
        foreach ($result->rows as $row) {
            $worksheet->setRow( $i, 26 );
            $price = $row['price'];
            $worksheet->writeString( $i, $j++, $row['sku'] );
            $worksheet->write( $i, $j++, $price );
            $worksheet->write( $i, $j++, $row['quantity'] );
            $i += 1;
            $j = 0;
        }
    }


    function populateSpecialsWorksheet( &$worksheet, &$database, &$priceFormat, &$boxFormat, &$textFormat )
    {
       // Set the column widths
       $j = 0;
       $worksheet->setColumn($j,$j++,strlen('sku')+1);
       $worksheet->setColumn($j,$j++,strlen('customer_group')+1);
       $worksheet->setColumn($j,$j++,strlen('priority')+1);
       $worksheet->setColumn($j,$j++,max(strlen('price'),10)+1,$priceFormat);
       $worksheet->setColumn($j,$j++,max(strlen('date_start'),19)+1,$textFormat);
       $worksheet->setColumn($j,$j++,max(strlen('date_end'),19)+1,$textFormat);

       // The heading row
       $i = 0;
       $j = 0;
       $worksheet->writeString( $i, $j++, 'sku', $boxFormat );
       $worksheet->writeString( $i, $j++, 'customer_group', $boxFormat );
       $worksheet->writeString( $i, $j++, 'priority', $boxFormat );
       $worksheet->writeString( $i, $j++, 'price', $boxFormat );
       $worksheet->writeString( $i, $j++, 'date_start', $boxFormat );
       $worksheet->writeString( $i, $j++, 'date_end', $boxFormat );
       $worksheet->setRow( $i, 30, $boxFormat );

       // The actual product specials data
       $i += 1;
       $j = 0;
       $query  = "SELECT p.sku, ps.*, cgd.name FROM `".DB_PREFIX."product_special` ps ";
       $query .= "INNER JOIN `".DB_PREFIX."product` p ON p.product_id=ps.product_id ";
       $query .= "LEFT JOIN `".DB_PREFIX."customer_group` cg ON cg.customer_group_id=ps.customer_group_id ";
       $query .= "LEFT JOIN `".DB_PREFIX."customer_group_description` cgd ON cg.customer_group_id=cgd.customer_group_id AND language_id = '" . (int) $this->config->get('config_language_id') . "' ";
       $query .= "ORDER BY ps.product_id, cgd.name";
       $result = $database->query( $query );
       foreach ($result->rows as $row) {
           $worksheet->setRow( $i, 13 );
           $worksheet->write( $i, $j++, $row['sku'] );
           $worksheet->write( $i, $j++, $row['name'] );
           $worksheet->write( $i, $j++, $row['priority'] );
           $worksheet->write( $i, $j++, $row['price'], $priceFormat );
           $worksheet->write( $i, $j++, $row['date_start'], $textFormat );
           $worksheet->write( $i, $j++, $row['date_end'], $textFormat );
           $i += 1;
           $j = 0;
       }
       }


       function populateDiscountsWorksheet( &$worksheet, &$database, &$priceFormat, &$boxFormat, &$textFormat )
       {
           // Set the column widths
           $j = 0;
           $worksheet->setColumn($j,$j++,strlen('sku')+1);
           $worksheet->setColumn($j,$j++,strlen('customer_group')+1);
           $worksheet->setColumn($j,$j++,strlen('quantity')+1);
           $worksheet->setColumn($j,$j++,strlen('priority')+1);
           $worksheet->setColumn($j,$j++,max(strlen('price'),10)+1,$priceFormat);
           $worksheet->setColumn($j,$j++,max(strlen('date_start'),19)+1,$textFormat);
           $worksheet->setColumn($j,$j++,max(strlen('date_end'),19)+1,$textFormat);

           // The heading row
           $i = 0;
           $j = 0;
           $worksheet->writeString( $i, $j++, 'sku', $boxFormat );
           $worksheet->writeString( $i, $j++, 'customer_group', $boxFormat );
           $worksheet->writeString( $i, $j++, 'quantity', $boxFormat );
           $worksheet->writeString( $i, $j++, 'priority', $boxFormat );
           $worksheet->writeString( $i, $j++, 'price', $boxFormat );
           $worksheet->writeString( $i, $j++, 'date_start', $boxFormat );
           $worksheet->writeString( $i, $j++, 'date_end', $boxFormat );
           $worksheet->setRow( $i, 30, $boxFormat );

           // The actual product discounts data
           $i += 1;
           $j = 0;
           $query  = "SELECT p.sku, pd.*, cgd.name FROM `".DB_PREFIX."product_discount` pd ";
           $query .= "INNER JOIN `".DB_PREFIX."product` p ON p.product_id=pd.product_id ";
           $query .= "LEFT JOIN `".DB_PREFIX."customer_group` cg ON cg.customer_group_id=pd.customer_group_id ";
           $query .= "LEFT JOIN `".DB_PREFIX."customer_group_description` cgd ON cg.customer_group_id=cgd.customer_group_id AND language_id = '" . (int) $this->config->get('config_language_id') . "' ";
           $query .= "ORDER BY pd.product_id, cgd.name";
           $result = $database->query( $query );
           foreach ($result->rows as $row) {
               $worksheet->setRow( $i, 13 );
               $worksheet->write( $i, $j++, $row['sku'] );
               $worksheet->write( $i, $j++, $row['name'] );
               $worksheet->write( $i, $j++, $row['quantity'] );
               $worksheet->write( $i, $j++, $row['priority'] );
               $worksheet->write( $i, $j++, $row['price'], $priceFormat );
               $worksheet->write( $i, $j++, $row['date_start'], $textFormat );
               $worksheet->write( $i, $j++, $row['date_end'], $textFormat );
               $i += 1;
               $j = 0;
           }
       }


    protected function clearSpreadsheetCache()
    {
        $files = glob(DIR_CACHE . 'Spreadsheet_Excel_Writer' . '*');

        if ($files) {
            foreach ($files as $file) {
                if (file_exists($file)) {
                    @unlink($file);
                    clearstatcache();
                }
            }
        }
    }


    function download()
    {
        global $config;
        global $log;
        $config = $this->config;
        $log = $this->log;
        set_error_handler('error_handler_for_export',E_ALL);
        register_shutdown_function('fatal_error_shutdown_handler_for_export');
        $database =& $this->db;

        // We use the package from http://pear.php.net/package/Spreadsheet_Excel_Writer/
        chdir( '../system/pear' );
        require_once "Spreadsheet/Excel/Writer.php";
        chdir( '../../admin' );

        // Creating a workbook
        $workbook = new Spreadsheet_Excel_Writer();
        $workbook->setTempDir(DIR_CACHE);
        $workbook->setVersion(8); // Use Excel97/2000 BIFF8 Format
        $boxFormat =& $workbook->addFormat(array('Size' => 10,'vAlign' => 'vequal_space' ));
        $textFormat =& $workbook->addFormat(array('Size' => 10, 'NumFormat' => "@" ));

        // sending HTTP headers
        $workbook->send('products.xls');

        // Creating the products worksheet
        $worksheet =& $workbook->addWorksheet('Products');
        $worksheet->setInputEncoding ( 'UTF-8' );
        $this->populateProductsWorksheet( $worksheet, $database, $boxFormat, $textFormat );
        $worksheet->freezePanes(array(1, 1, 1, 1));

        // Creating the specials worksheet
        $worksheet =& $workbook->addWorksheet('Specials');
        $worksheet->setInputEncoding ( 'UTF-8' );
        $this->populateSpecialsWorksheet( $worksheet, $database, $priceFormat, $boxFormat, $textFormat );
        $worksheet->freezePanes(array(1, 1, 1, 1));

        // Creating the discounts worksheet
        $worksheet =& $workbook->addWorksheet('Discounts');
        $worksheet->setInputEncoding ( 'UTF-8' );
        $this->populateDiscountsWorksheet( $worksheet, $database, $priceFormat, $boxFormat, $textFormat );
        $worksheet->freezePanes(array(1, 1, 1, 1));

        // Let's send the file
        $workbook->close();

        // Clear the spreadsheet caches
        $this->clearSpreadsheetCache();
        exit;
    }
}
