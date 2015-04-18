<?php

/** Database Connection file * */
include 'conn.php';

/** PHPExcel */
include 'PHPExcel_1.7.9_doc/Classes/PHPExcel.php';

/** PHPExcel_Writer_Excel2007 */
include 'PHPExcel_1.7.9_doc/Classes/PHPExcel/Writer/Excel2007.php';

ini_set('display_startup_errors',1);
ini_set('display_errors',1);
error_reporting(-1);
// Create new PHPExcel object
//echo date('H:i:s') . " Report Generated<br>";
$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Keshav");
//$objPHPExcel->getProperties()->setTitle("Boneyard Report");
//$objPHPExcel->getProperties()->setSubject("Boneyard Report");
//$objPHPExcel->getProperties()->setDescription("Boneyard Report");
$objPHPExcel->setActiveSheetIndex(0);
$rowCount = 2;

if ($_POST) {
    $date['dateFrom'] = $_POST['dateFrom'];
    $date['dateTo'] = $_POST['dateTo'];
    
    $time['timeFrom'] = isset($_POST['timeFrom']) ? $_POST['timeFrom'] : "";
    $time['timeTo'] = isset($_POST['timeTo']) ? $_POST['timeTo'] : "";
    
    $opt = $_POST['opt'];        
    
    /** Getting order ID's according to the date * */
    $order_id_arr = getOrderIDs($date);
//    if (empty($order_id_arr)) {
//        header("Location: form.php?error=nodata");
//    } else {
        /** Getting product ids on the basis of order ids * */
        $product_id_arr = getProductIDs($order_id_arr);
        /** Getting filtered data for chargeable and non chargeable fields * */
        $filtered_product_ids['to_c'] = getFilterDataForBoneyard($date, 'OUT', 'chargeable');
        $filtered_product_ids['to_nc'] = getFilterDataForBoneyard($date, 'OUT', 'non chargeable');

        $filtered_product_ids['from_c'] = getFilterDataForBoneyard($date, 'IN', 'chargeable');
        $filtered_product_ids['from_nc'] = getFilterDataForBoneyard($date, 'IN', 'non chargeable');

        /** Getting data from Total table * */
        $total_data = getTotalData($product_id_arr,$date);
        if($opt == "boneyard") {
            $fName = "Boneyard Report";
        } 
        else if($opt == "pandl"){
            $fName = "P and L Report";
        }
        else if($opt == "spandl") {
            $fName = "Stand P and L Report";
        }
        else if($opt == "orderrep") {
            $fName = "Order Report";
        }
        else if($opt == "prodrep") {
            $fName = "Product Report";
        } else if($opt == "voidrep") {
            $fName = "VOID Report";
        } else if($opt == "locordrep") {
            $fName = "Location Order Report";
        } else if($opt == "prodlist") {
            $fName = "Product List";
        } else if($opt == "loclist") {
            $fName = "Location List";
        } else if($opt == "boneyardcountdown") {
            $fName = "Boneyard Count Down";
        }

        // We'll be outputting an excel file
        header('Content-type: application/vnd.ms-excel');

        // Give default name to the file
        if($opt == "loclist" || $opt == "prodlist") {
            header('Content-Disposition: attachment; filename="'.$fName.'"' . ".xls");
        } else {
            header('Content-Disposition: attachment; filename="'.$fName.'"' . $date['dateFrom'] ."through". $date['dateTo'] . ".xls");
        }
        
        if($opt == "boneyard") {
            $total_received_data = getTotalReceivedData($date);
            $total_spoilage_data = getSpoilageTotalData(array(),$date,array());
            //var_dump($total_received_data);
            getProductInfo($product_id_arr, $filtered_product_ids, $total_data,$total_received_data,$total_spoilage_data);
        }            
        else if($opt == "pandl"){
            $rowCount = 3;
            $time['timeFrom'] = "00:01";
            $time['timeTo'] = "23:59";
            $chargeable_loc_data = getChargeableLocation();
            $fromTo = "NCToC";
            $transfers[$fromTo] = getProductCToNCAndNCToC($chargeable_loc_data,"NCToC",$date,$time);
            $fromTo = "CToNC";
            $transfers[$fromTo] = getProductCToNCAndNCToC($chargeable_loc_data,"CToNC",$date,$time);
            $total_spoilage_data = getSpoilageTotalData($product_id_arr,$date,$time);
            $total_client_data = getClientFieldInventory($date,$time);
            $total_sales_data = getTotalSales($date,$time);
            $total_beg_field_inv = getBegFieldInv();
            $total_received_data = getTotalReceivedData($date);
            $subcomm_data = getTotalData($order_id_arr,$date,'%sub comm%');
            if($date['dateFrom'] == $date['dateTo']) {
                $date1['dateFrom'] = date('Y-m-d', strtotime($date['dateFrom'] .' -1 day'));
                $check = 1;
            } else {
                $date1['dateFrom'] = $date['dateFrom'];
                $check = 0;
            }
            $date1['dateTo'] = date('Y-m-d', strtotime($date['dateTo'] .' -1 day'));
            //$order_id_arr1 = getOrderIDs($date1);
            $prior_total_data = ($check == 0) ? array(): getTotalData(array(),$date1);
            $prior_subcomm_data = getTotalData(array(),$date1,'%sub comm%');
            $prior_client_data = getClientFieldInventory($date1,$time);
            
            getPAndLInfo($product_id_arr,$filtered_product_ids,$transfers,$total_data,$total_spoilage_data,$total_client_data,$total_sales_data,$total_beg_field_inv,$total_received_data,$subcomm_data,$prior_total_data,$prior_subcomm_data,$prior_client_data);
        }
        else if($opt == "spandl") {
            $rowCount = 6;
            $time['timeFrom'] = "00:01";
            $time['timeTo'] = "23:59";
            $location = $_POST['location'];
            $chargeable_loc_data = getChargeableLocation();
            $fromTo = "NCToC";
            $transfers[$fromTo] = getProductCToNCAndNCToC($chargeable_loc_data,"NCToC",$date,$time,$location);
            $fromTo = "CToNC";
            $transfers[$fromTo] = getProductCToNCAndNCToC($chargeable_loc_data,"CToNC",$date,$time,$location);
            $fromTo = "CToC";
            $transfers[$fromTo] = getProductCToNCAndNCToC($chargeable_loc_data,"CToC",$date,$time,$location);
            $fromTo = "PCToC";
            $transfers[$fromTo] = getProductCToNCAndNCToC($chargeable_loc_data,"PCToC",$date,$time,$location);
            $total_spoilage_data = getSpoilageTotalData($product_id_arr,$date,$time,$location);
            $total_client_data = getClientFieldInventory($date,$time,$location);
            $total_sales_data = getTotalSales($date,$time,$location);
            $total_beg_field_inv = getBegFieldInv();
            $standBegInv = standBegFieldInv($location,$date);                
            $date1['dateFrom'] = date('Y-m-d', strtotime($date['dateTo'] .' -1 day'));
            $date1['dateTo'] = date('Y-m-d', strtotime($date['dateTo'] .' -1 day'));

            $prior_cfe_data = getClientFieldInventory($date1,$time,$location);
            $val = array();
            if($date['dateFrom'] == $date['dateTo']) {                
                $date1['dateFrom'] = date('Y-m-d', strtotime($date['dateFrom'] .' -1 day'));
                $date1['dateTo'] = date('Y-m-d', strtotime($date['dateFrom'] .' -1 day'));
                $order_id_arr1 = getOrderIDs($date1);
                $product_id_arr1 = getProductIDs($order_id_arr1);
                $filtered_product_ids1['to_c'] = getFilterDataForBoneyard($date1, 'OUT', 'chargeable');
                $filtered_product_ids1['to_nc'] = getFilterDataForBoneyard($date1, 'OUT', 'non chargeable');

                $filtered_product_ids1['from_c'] = getFilterDataForBoneyard($date1, 'IN', 'chargeable');
                $filtered_product_ids1['from_nc'] = getFilterDataForBoneyard($date1, 'IN', 'non chargeable');
                
                $total_data1 = getTotalData($product_id_arr,$date1);
                $fromTo = "NCToC";
                $transfers1[$fromTo] = getProductCToNCAndNCToC($chargeable_loc_data,"NCToC",$date1,$time,$location);
                $fromTo = "CToNC";
                $transfers1[$fromTo] = getProductCToNCAndNCToC($chargeable_loc_data,"CToNC",$date1,$time,$location);
                $fromTo = "CToC";
                $transfers1[$fromTo] = getProductCToNCAndNCToC($chargeable_loc_data,"CToC",$date1,$time,$location);
                $fromTo = "PCToC";
                $transfers[$fromTo] = getProductCToNCAndNCToC($chargeable_loc_data,"PCToC",$date,$time,$location);
                $total_spoilage_data1 = getSpoilageTotalData($product_id_arr,$date1,$time,$location);
                $total_client_data1 = getClientFieldInventory($date1,$time,$location);
                $total_sales_data1 = getTotalSales($date1,$time,$location);            
                $standBegInv1 = standBegFieldInv($location,$date1);
                $val = getPreviousDayValue($product_id_arr1,$filtered_product_ids1,$transfers1,$total_data1,$total_spoilage_data1,$total_client_data1,$total_sales_data1,$total_beg_field_inv,$standBegInv1);
            }
//            var_dump($val);
            getStandPAndLInfo($product_id_arr,$filtered_product_ids,$transfers,$total_data,$total_spoilage_data,$total_client_data,$total_sales_data,$total_beg_field_inv,$standBegInv,$val,$prior_cfe_data);
        }
        else if($opt == "orderrep") {
            getOrderReports($date);
        }
        else if($opt == "prodrep") {
            getProductReport($date);
        } else if($opt == "voidrep") {
            getOrderReports($date,"void");
        } else if($opt == "locordrep") {
            $location = $_POST['location1'];
            getOrderReports($date,"location",$location);
        } else if($opt == "prodlist") {
            getProductList();
        } elseif($opt == "loclist") {
            getLocationList();
        } else if($opt == "boneyardcountdown") {
            getProductListBC();
        }
    //}
} else {
    echo '<b>Invalid access to file export.php</b>';
}

/** Getting order id's on the basis of date * */
function getOrderIDs($date) {
    global $conn;
    $stmt = $conn->prepare('SELECT ORDER_ID FROM `order` WHERE DATE(`date`) between :dt1 and :dt2');
    $stmt->execute(array('dt1' => $date['dateFrom'], 'dt2' => $date['dateTo']));
    $order_id_arr = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $order_id_arr[] = $row['ORDER_ID'];
    }
    return $order_id_arr;
}

/** Getting data on the basis of type, location and chargeable * */
function getFilterDataForBoneyard($date, $type, $chargeable) {
    global $conn;
    if ($type == 'IN') {
        $loc1 = 'LOCATION_TO';
        $loc2 = 'LOCATION_FROM';
    } else {
        $loc2 = 'LOCATION_TO';
        $loc1 = 'LOCATION_FROM';
    }
    $stmt = $conn->prepare("SELECT PRODUCT_ID,SUM(QTY_PALLET) as QTY_PALLET,SUM(QTY_CASE) as QTY_CASE,SUM(QTY_BOTTLE) as QTY_BOTTLE,SUM(QTY_PARTIAL) as QTY_PARTIAL FROM `bill` WHERE ORDER_ID IN (SELECT ORDER_ID FROM `order` WHERE DATE(`date`) between :dt1 and :dt2 AND TYPE = :type AND $loc1 = :loc AND $loc2 IN (SELECT NAME FROM location WHERE CHARGEABLE = :charge) AND STATUS != :status ) GROUP BY PRODUCT_ID");
    $stmt->execute(array('dt1' => $date['dateFrom'], 'dt2' => $date['dateTo'], 'type' => $type, 'loc' => 'boneyard', 'charge' => $chargeable, 'status' => 'VOID'));
    $product_id_boneyard_arr = array();
    $qty_pallet_arr = array();
    $qty_case_arr = array();
    $qty_bottle_arr = array();
    $qty_partial_arr = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $product_id_boneyard_arr [] = $row['PRODUCT_ID'];
        $qty_pallet_arr[$row['PRODUCT_ID']] = $row['QTY_PALLET'];
        $qty_case_arr[$row['PRODUCT_ID']] = $row['QTY_CASE'];
        $qty_bottle_arr[$row['PRODUCT_ID']] = $row['QTY_BOTTLE'];
        $qty_partial_arr[$row['PRODUCT_ID']] = $row['QTY_PARTIAL'];
    }
    $filter_data['PRODUCT_ID'] = $product_id_boneyard_arr;
    $filter_data['QTY_PALLET'] = $qty_pallet_arr;
    $filter_data['QTY_CASE'] = $qty_case_arr;
    $filter_data['QTY_BOTTLE'] = $qty_bottle_arr;
    $filter_data['QTY_PARTIAL'] = $qty_partial_arr;
    return $filter_data;
}

/** Getting product ids from Bill table on the basis of order id's * */
function getProductIDs($order_id_arr) {
    global $conn;
    $stmt = $conn->prepare('SELECT PRODUCT_ID FROM `bill` WHERE ORDER_ID IN (0' . implode(",", $order_id_arr) . ')');
    $stmt->execute();
    $product_id_arr = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $product_id_arr[] = $row['PRODUCT_ID'];
    }

    return $product_id_arr;
}

/** Getting data from Table data * */
function getTotalData($product_id_arr,$date1,$location="Boneyard") {
    global $conn;
    $stmt = $conn->prepare("SELECT PRODUCT,ENDING_PALLET,ENDING_CASE,ENDING_BOTTLE,ENDING_PARTIAL FROM `total` WHERE LOCATION LIKE '$location' AND DATE(`date`) BETWEEN :dt1 and :dt2");
    $stmt->execute(array('dt1'=>$date1['dateFrom'],'dt2'=>$date1['dateTo']));
    $total_data = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $total_data['ENDING_PALLET'][strtolower($row['PRODUCT'])] = $row['ENDING_PALLET'];
        $total_data['ENDING_CASE'][strtolower($row['PRODUCT'])] = $row['ENDING_CASE'];
        $total_data['ENDING_BOTTLE'][strtolower($row['PRODUCT'])] = $row['ENDING_BOTTLE'];
        $total_data['ENDING_PARTIAL'][strtolower($row['PRODUCT'])] = $row['ENDING_PARTIAL'];
    }
    //var_dump($date1);
    //var_dump($total_data);
    //exit();
    return $total_data;
}

function getProductList()
{
    global $conn, $objPHPExcel;
    $rowCount = 2;
    /** Writing all the heading values * */
    // $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Total Venue Control LLC.');
    // $objPHPExcel->getActiveSheet()->SetCellValue('A2', 'Boneyard Report');
    $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Product Name');
    //$objPHPExcel->getActiveSheet()->SetCellValue('B1', 'Chargeable/Non Chargeable');

    $stmt = $conn->prepare('SELECT name FROM `product`');
    $stmt->execute();
    $styleArray = array(
        'borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_MEDIUM
            )
        )
    );

    $objPHPExcel->getActiveSheet()->getStyle("A1")->applyFromArray($styleArray);
    //$objPHPExcel->getActiveSheet()->getStyle("B1")->applyFromArray($styleArray);
    //$product_id_arr = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $objPHPExcel->getActiveSheet()->SetCellValue('A' . $rowCount, $row['name']);
      //  $objPHPExcel->getActiveSheet()->SetCellValue('B' . $rowCount, $row['chargeable']);
        $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray($styleArray);
        $rowCount++;
    }

    $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
    //$objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'EEECE1')
            )
        )
    );

    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(34.14);
    //$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(34.14);

    /** Saving the file * */
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    ob_end_clean();
    $objWriter->save('php://output');

}


function getProductListBC()
{
    global $conn, $objPHPExcel,$date;
    $rowCount = 4;
    /** Writing all the heading values * */
     $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Total Venue Control LLC.');
     $objPHPExcel->getActiveSheet()->SetCellValue('A2', 'Boneyard Count Down Report');
     $objPHPExcel->getActiveSheet()->SetCellValue('A3', "Date: {$date['dateFrom']} thru {$date['dateTo']}");

     //$objPHPExcel->getActiveSheet()->SetCellValue('A4', 'Product Name');

    $objPHPExcel->getActiveSheet()->SetCellValue('B1', 'Client Signature');
    $objPHPExcel->getActiveSheet()->SetCellValue('B3', 'TVC Signature');

    $stmt = $conn->prepare('SELECT name FROM `product`');
    $stmt->execute();
    $styleArray = array(
        'borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_MEDIUM
            )
        )
    );

    //$objPHPExcel->getActiveSheet()->getStyle("A1:A3")->applyFromArray($styleArray);
    //$objPHPExcel->getActiveSheet()->getStyle("B1")->applyFromArray($styleArray);
    //$product_id_arr = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $objPHPExcel->getActiveSheet()->SetCellValue('A' . $rowCount, $row['name']);
        //  $objPHPExcel->getActiveSheet()->SetCellValue('B' . $rowCount, $row['chargeable']);
        $objPHPExcel->getActiveSheet()->getStyle("A$rowCount:B$rowCount")->applyFromArray($styleArray);
        $objPHPExcel->getActiveSheet()->getRowDimension($rowCount)->setRowHeight(33.75);

        $rowCount++;
    }

    $objPHPExcel->getActiveSheet()->getStyle('A1:A3')->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle('B1:B3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
//    $objPHPExcel->getActiveSheet()->getStyle('A1:A3')->applyFromArray(
//        array(
//            'fill' => array(
//                'type' => PHPExcel_Style_Fill::FILL_SOLID,
//                'color' => array('rgb' => 'EEECE1')
//            )
//        )
//    );

    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(34.14);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(152.89);

    /** Saving the file * */
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    ob_end_clean();
    $objWriter->save('php://output');

}



function getLocationList()
{
    global $conn, $objPHPExcel;
    $rowCount = 2;
    /** Writing all the heading values * */
   // $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Total Venue Control LLC.');
   // $objPHPExcel->getActiveSheet()->SetCellValue('A2', 'Boneyard Report');
    $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Location Name');
    $objPHPExcel->getActiveSheet()->SetCellValue('B1', 'Chargeable/Non Chargeable');

    $stmt = $conn->prepare('SELECT name,chargeable FROM `location`');
    $stmt->execute();
    $styleArray = array(
        'borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_MEDIUM
            )
        )
    );

    $objPHPExcel->getActiveSheet()->getStyle("A1")->applyFromArray($styleArray);
    $objPHPExcel->getActiveSheet()->getStyle("B1")->applyFromArray($styleArray);
    //$product_id_arr = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $objPHPExcel->getActiveSheet()->SetCellValue('A' . $rowCount, $row['name']);
        $objPHPExcel->getActiveSheet()->SetCellValue('B' . $rowCount, $row['chargeable']);
        $objPHPExcel->getActiveSheet()->getStyle("A$rowCount:B$rowCount")->applyFromArray($styleArray);
        $rowCount++;
    }

    $objPHPExcel->getActiveSheet()->getStyle('A1:B1')->getFont()->setBold(true);
    //$objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle('A1:B1')->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'EEECE1')
            )
        )
    );

    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(34.14);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(34.14);

    /** Saving the file * */
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    ob_end_clean();
    $objWriter->save('php://output');

}

function getSpoilageTotalData( $product_id_arr,$date,$time,$location = "" ) {
    
    global $conn;
    $locate = ($location == "") ? "" : "AND LOCATION = '".$location. "'";
    if(empty($time)) {
        $stmt = $conn->prepare("SELECT PRODUCT,SUM(SPOILAGE) AS SPOILAGE FROM `total` WHERE `date` BETWEEN :dt1 and :dt2 AND LOCATION = 'Boneyard' $locate GROUP BY PRODUCT");
        $stmt->execute(array('dt1' => $date['dateFrom'], 'dt2' => $date['dateTo']));
    } else {
        $stmt = $conn->prepare("SELECT PRODUCT,SUM(SPOILAGE) AS SPOILAGE FROM `total` WHERE `date` BETWEEN :dt1 and :dt2 AND `time` BETWEEN :tm1 AND :tm2 $locate GROUP BY PRODUCT");
        $stmt->execute(array('dt1' => $date['dateFrom'],"tm1"=>$time['timeFrom'], 'dt2' => $date['dateTo'],"tm2"=>$time['timeTo']));
    }
    $total_spoilage_data = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $total_spoilage_data[$row['PRODUCT']] = $row['SPOILAGE'];
    }
    return $total_spoilage_data;
}

function getChargeableLocation()
{
    global $conn;
    $stmt = $conn->prepare('SELECT NAME,CHARGEABLE FROM location WHERE NAME != :name');
    $stmt->execute(array('name'=>'Boneyard'));
    $chargeable_loc_data = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        if($row['CHARGEABLE'] == 'Chargeable') {
            $chargeable_loc_data['CHARGEABLE'][] = $row['NAME'];
        }
        else {
            $chargeable_loc_data['NONCHARGEABLE'][] = $row['NAME'];
        }
    }
    
    return $chargeable_loc_data;
}

function getTotalReceivedData($date,$less = "")
{
    global $conn;
    if($less == "") {
        $stmt = $conn->prepare('SELECT `product`,SUM(`total_pallet`) as total_pallet,SUM(`total_case`) as total_case,SUM(`total_individual`) as total_individual,SUM(`total_partial`) as total_partial FROM `totaladddate` WHERE `date` BETWEEN :dt1 and :dt2 GROUP BY `product`');
        $stmt->execute(array('dt1'=>$date['dateFrom'],'dt2'=>$date['dateTo']));
    } else {
        $stmt = $conn->prepare('SELECT `product`,SUM(`total_pallet`) as total_pallet,SUM(`total_case`) as total_case,SUM(`total_individual`) as total_individual,SUM(`total_partial`) as total_partial FROM `totaladddate` WHERE `date` <= :dt1 GROUP BY `product`');
        $stmt->execute(array('dt1'=>$date['dateFrom']));
    }
    
    $total_received_data = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $total_received_data[$row['product']]['total_pallet'] = $row['total_pallet'];
        $total_received_data[$row['product']]['total_case'] = $row['total_case'];
        $total_received_data[$row['product']]['total_individual'] = $row['total_individual'];
        $total_received_data[$row['product']]['total_partial'] = $row['total_partial'];
    }
    
    return $total_received_data;
    
}

/** Generating excel sheet and calculating all the values * */
function getProductInfo($product_ids, $filtered_product_ids, $total_data,$total_received_data,$total_spoilage_data) {
    global $conn, $objPHPExcel, $date;
    $rowCount = 6;
    $stmt = $conn->prepare('SELECT ID,NAME,SIZE,QTY_PER_CASE,QTY_PER_PALLET,OPENING_PALLET,OPENING_CASE,OPENING_BOTTLE,OPENING_PARTIAL,CUR_PALLET,CUR_CASE,CUR_BOTTLE,CUR_PARTIAL FROM `product`');
    $stmt->execute();
    
    if($date['dateFrom'] == $date['dateTo']) {
        $d_name = 'Total Received Inventory';
    } else {
        $d_name = 'Total Received Inventory';
    }
    
    /** Writing all the heading values * */
    $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Total Venue Control LLC.');
    $objPHPExcel->getActiveSheet()->SetCellValue('A2', 'Boneyard Report');
    $objPHPExcel->getActiveSheet()->SetCellValue('A3', "Date: {$date['dateFrom']} thru {$date['dateTo']}");
    
    $objPHPExcel->getActiveSheet()->SetCellValue('A4', 'Product Description');
    //$objPHPExcel->getActiveSheet()->SetCellValue('B1', 'Size');
    //$objPHPExcel->getActiveSheet()->SetCellValue('B4', 'Case Count');
    //$objPHPExcel->getActiveSheet()->SetCellValue('C4', 'Pallet Count');
    $objPHPExcel->getActiveSheet()->SetCellValue('B4', "Today's Receiving");
    $objPHPExcel->getActiveSheet()->SetCellValue('C4', 'Prior Day Ending Inventory');    
    $objPHPExcel->getActiveSheet()->SetCellValue('D4', $d_name);
    $objPHPExcel->getActiveSheet()->SetCellValue('E4', 'To "C" Stands');
    $objPHPExcel->getActiveSheet()->SetCellValue('F4', 'To "NC" Stands');
    $objPHPExcel->getActiveSheet()->SetCellValue('G4', 'From "C" Stands');
    $objPHPExcel->getActiveSheet()->SetCellValue('H4', 'From "NC" Stands');
    $objPHPExcel->getActiveSheet()->SetCellValue('I4', 'Spoilage');
    $objPHPExcel->getActiveSheet()->SetCellValue('J4', 'Perpetual Ending Inventory Per Unit');
    $objPHPExcel->getActiveSheet()->SetCellValue('K4', 'Actual Ending Inventory Per Unit');
    $objPHPExcel->getActiveSheet()->SetCellValue('L4', 'Difference Actual To Perpetual Per Unit');
    $objPHPExcel->getActiveSheet()->SetCellValue('M4', 'Difference Actual To Perpetual Per Case');
    
    $styleArray = array(
      'borders' => array(
          'allborders' => array(
              'style' => PHPExcel_Style_Border::BORDER_MEDIUM
          )
      )
    );
    $p_total_data = array();
    $check = 0;
    if($date['dateFrom'] == $date['dateTo']) {
        $date1['dateFrom'] = date('Y-m-d', strtotime($date['dateFrom'] .' -1 day'));
        $date1['dateTo'] = date('Y-m-d', strtotime($date['dateFrom'] .' -1 day'));
        //echo "date : ";
        //var_dump($date1);
        //$p_total_data = getTotalReceivedData($date1,"less");
        $p_total_data = getTotalData(array(),$date1);
        $check = 1;
        //echo "if";
    }
    /** Fetching all the data * */
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        //$total_received = ($row['QTY_PER_PALLET'] * $row['OPENING_PALLET'] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $row['OPENING_CASE']) + $row['OPENING_BOTTLE'] + $row['OPENING_PARTIAL'];
        $total_received = ($row['QTY_PER_PALLET'] * $total_received_data[$row['NAME']]['total_pallet'] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $total_received_data[$row['NAME']]['total_case']) + $total_received_data[$row['NAME']]['total_individual'] + $total_received_data[$row['NAME']]['total_partial'];
        $p_aei = 0;
        if(!empty($p_total_data)) {
            $p_aei = (($row['QTY_PER_PALLET'] * $p_total_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $p_total_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($p_total_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $p_total_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
            //$total_received = $total_received + $p_aei;
        }
        
        $to_c = 0;
        $to_nc = 0;
        $from_c = 0;
        $from_nc = 0;
        $pei = 0;
        $aei = 0;
        $diff = 0;
        if (in_array($row['ID'], $product_ids)) {
            if (in_array($row['ID'], $filtered_product_ids['to_c']['PRODUCT_ID'])) {
                $to_c = (($row['QTY_PER_PALLET'] * $filtered_product_ids['to_c']['QTY_PALLET'][$row['ID']] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $filtered_product_ids['to_c']['QTY_CASE'][$row['ID']]) + ($filtered_product_ids['to_c']['QTY_BOTTLE'][$row['ID']] + $filtered_product_ids['to_c']['QTY_PARTIAL'][$row['ID']]));
            }

            if (in_array($row['ID'], $filtered_product_ids['to_nc']['PRODUCT_ID'])) {
                $to_nc = (($row['QTY_PER_PALLET'] * $filtered_product_ids['to_nc']['QTY_PALLET'][$row['ID']] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $filtered_product_ids['to_nc']['QTY_CASE'][$row['ID']]) + ($filtered_product_ids['to_nc']['QTY_BOTTLE'][$row['ID']] + $filtered_product_ids['to_nc']['QTY_PARTIAL'][$row['ID']]));
            }

            if (in_array($row['ID'], $filtered_product_ids['from_c']['PRODUCT_ID'])) {
                $from_c = (($row['QTY_PER_PALLET'] * $filtered_product_ids['from_c']['QTY_PALLET'][$row['ID']] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $filtered_product_ids['from_c']['QTY_CASE'][$row['ID']]) + ($filtered_product_ids['from_c']['QTY_BOTTLE'][$row['ID']] + $filtered_product_ids['from_c']['QTY_PARTIAL'][$row['ID']]));
            }

            if (in_array($row['ID'], $filtered_product_ids['from_nc']['PRODUCT_ID'])) {
                $from_nc = (($row['QTY_PER_PALLET'] * $filtered_product_ids['from_nc']['QTY_PALLET'][$row['ID']] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $filtered_product_ids['from_nc']['QTY_CASE'][$row['ID']]) + ($filtered_product_ids['from_nc']['QTY_BOTTLE'][$row['ID']] + $filtered_product_ids['from_nc']['QTY_PARTIAL'][$row['ID']]));
            }

//            $pei = ( $total_received - ($to_c + $to_nc) ) + $from_c + $from_nc;
//            $aei = (($row['QTY_PER_PALLET'] * $total_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $total_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($total_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $total_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
//            $diff = $pei - $aei;
        }
             $pei = ( $total_received+$p_aei - ($to_c + $to_nc) ) + $from_c + $from_nc;
            $aei = (($row['QTY_PER_PALLET'] * $total_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $total_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($total_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $total_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
            $diff = $pei - $aei;
            
            $b_row = ($check == 1)?$total_received:0;
            $d_row = ($check == 1)?"=SUM(B$rowCount:C$rowCount)":$total_received;
        /** Setting all the values in the report * */
        $objPHPExcel->getActiveSheet()->SetCellValue('A' . $rowCount, $row['NAME']);
        //$objPHPExcel->getActiveSheet()->SetCellValue('B' . $rowCount, $row['SIZE']);
        //$objPHPExcel->getActiveSheet()->SetCellValue('B' . $rowCount, $row['QTY_PER_CASE']);
        //$objPHPExcel->getActiveSheet()->SetCellValue('C' . $rowCount, $row['QTY_PER_PALLET']);
        $objPHPExcel->getActiveSheet()->SetCellValue('B' . $rowCount, $b_row);
        $objPHPExcel->getActiveSheet()->SetCellValue('C' . $rowCount, $p_aei);
        $objPHPExcel->getActiveSheet()->SetCellValue('D' . $rowCount, $d_row);
        $objPHPExcel->getActiveSheet()->SetCellValue('E' . $rowCount, -$to_c);
        $objPHPExcel->getActiveSheet()->SetCellValue('F' . $rowCount, -$to_nc);
        $objPHPExcel->getActiveSheet()->SetCellValue('G' . $rowCount, $from_c);
        $objPHPExcel->getActiveSheet()->SetCellValue('H' . $rowCount, $from_nc);
        $objPHPExcel->getActiveSheet()->SetCellValue('I' . $rowCount, -$total_spoilage_data[$row['NAME']]);
        $objPHPExcel->getActiveSheet()->SetCellValue('J' . $rowCount, "=SUM(D$rowCount:I$rowCount)");
        $objPHPExcel->getActiveSheet()->SetCellValue('K' . $rowCount, $aei);
        $objPHPExcel->getActiveSheet()->SetCellValue('L' . $rowCount, "=SUM(D$rowCount:I$rowCount)-K$rowCount");
        $objPHPExcel->getActiveSheet()->SetCellValue('M' . $rowCount, "=L$rowCount/{$row['QTY_PER_CASE']}");
        
        $objPHPExcel->getActiveSheet()->getStyle('A'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $objPHPExcel->getActiveSheet()->getStyle('B'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('C'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('D'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('E'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('F'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('G'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('H'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('I'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('J'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('K'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('L'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('M'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('N'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        
        $objPHPExcel->getActiveSheet()->getStyle( 'B'.($rowCount).':M'.$rowCount )->getNumberFormat()->setFormatCode(
            '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        );
        
        $objPHPExcel->getActiveSheet()->getStyle("A$rowCount:M$rowCount")->applyFromArray($styleArray);
        
        $rowCount++;
    }

    if($check == 0)
    {
        $objPHPExcel->getActiveSheet()->getColumnDimension("B")->setVisible(false);
        $objPHPExcel->getActiveSheet()->getColumnDimension("C")->setVisible(false);
    }

    $objPHPExcel->getActiveSheet()->getStyle('A1:A3')->getFont()->setBold(true);
    /** Formatting all the heading columns * */
    $objPHPExcel->getActiveSheet()->getStyle("A4:M4")->applyFromArray($styleArray);
    $objPHPExcel->getActiveSheet()->getStyle('A4:M4')->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle('A4:M4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle('A4:M4')->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'EEECE1')
            )
        )
    );
    
    $objPHPExcel->getActiveSheet()->getStyle('A5:M5')->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '00FF00')
            )
        )
    );
    
    $objPHPExcel->getActiveSheet()->getRowDimension('4')->setRowHeight(60.75);
    $objPHPExcel->getActiveSheet()->getRowDimension('5')->setRowHeight(4.5);
    $objPHPExcel->getActiveSheet()->getStyle('B4:M4')
    ->getAlignment()->setWrapText(true);
    
    
    
    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(30.00);
    //$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(12.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(12.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(12.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(12.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(11.14);
    $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(21.14);
    $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(21.14);
    $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(11.14);
    $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(11.14);
    //$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(11.14);

    /** Saving the file * */
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    ob_end_clean();
    $objWriter->save('php://output');
    //$objWriter->save('Boneyard Report.xlsx');
}

function getProductCToNCAndNCToC($chargeable_loc_data,$fromTo,$date,$time,$location = "")
{
    global $conn, $objPHPExcel, $rowCount;
    if($fromTo == "NCToC")
    {
        $loc_to = $chargeable_loc_data['CHARGEABLE'];
        $loc_from = $chargeable_loc_data['NONCHARGEABLE'];
        $locate = ($location == "") ? "" : "AND LOCATION_TO = '".$location. "'";
    } elseif($fromTo == "CToC") {
        $loc_to = $chargeable_loc_data['CHARGEABLE'];
        $loc_from = $chargeable_loc_data['CHARGEABLE'];
        $locate = ($location == "") ? "" : "AND LOCATION_FROM = '".$location. "'";
    } elseif($fromTo == "PCToC") {
        $loc_to = $chargeable_loc_data['CHARGEABLE'];
        $loc_from = $chargeable_loc_data['CHARGEABLE'];
        $locate = ($location == "") ? "" : "AND LOCATION_TO = '".$location. "'";
    }
    else
    {
        $loc_to = $chargeable_loc_data['NONCHARGEABLE'];
        $loc_from = $chargeable_loc_data['CHARGEABLE'];
        $locate = ($location == "") ? "" : "AND LOCATION_FROM = '".$location. "'";
    }
    
    //$stmt = $conn->prepare("SELECT PRODUCT_ID,COUNT(*) AS 'COUNT' FROM `BILL` WHERE ORDER_ID IN (SELECT ORDER_ID FROM `order` WHERE DATE(`date`) BETWEEN :dt1 and :dt2 AND LOCATION_FROM IN ($loc_from) AND LOCATION_TO IN ($loc_to) AND STATUS != :status) GROUP BY PRODUCT_ID");
    //$stmt->execute(array( 'dt1' => $date['dateFrom'] , 'dt2' => $date['dateTo'] , 'status' => 'VOID' ));
    //echo "SELECT PRODUCT_ID,COUNT(*) AS 'COUNT' FROM `BILL` WHERE ORDER_ID IN (SELECT ORDER_ID FROM `order` WHERE DATE(`date`) BETWEEN ? and ? AND TIME(`date`) BETWEEN ? AND ? AND LOCATION_FROM IN (".implode(',',array_fill(0, count($loc_from),'?')).") AND LOCATION_TO IN (".implode(',',array_fill(0, count($loc_to),'?')).") AND STATUS != ? $locate) GROUP BY PRODUCT_ID<br>";
    $stmt = $conn->prepare("SELECT `product_id`,SUM(`qty_pallet`) as `qty_pallet` ,SUM(`qty_case`) as `qty_case` ,SUM(`qty_bottle`) as `qty_bottle`,SUM(`qty_partial`) AS `qty_partial` FROM `bill` WHERE ORDER_ID IN (SELECT ORDER_ID FROM `order` WHERE `date` BETWEEN ? and ? AND LOCATION_FROM IN (".implode(',',array_fill(0, count($loc_from),'?')).") AND LOCATION_TO IN (".implode(',',array_fill(0, count($loc_to),'?')).") AND STATUS != ? $locate) GROUP BY PRODUCT_ID");
    $arr[] = $date['dateFrom']." ".$time['timeFrom'];
    $arr[] = $date['dateTo']." ".$time['timeTo'];
    //$arr[] = $time['timeFrom'];
    //$arr[] = $time['timeTo'];
    $arr1 = array_merge($arr,$loc_from,$loc_to);
    $arr1[] = 'VOID';    
    $stmt->execute( $arr1 );
      
    $transfers = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $transfers[$row['product_id']]['qty_pallet'] = $row['qty_pallet'];
        $transfers[$row['product_id']]['qty_case'] = $row['qty_case'];
        $transfers[$row['product_id']]['qty_bottle'] = $row['qty_bottle'];
        $transfers[$row['product_id']]['qty_partial'] = $row['qty_partial'];
    }
    
    return $transfers;
}

function getClientFieldInventory( $date,$time,$location = "" ) {
    
    global $conn;
    $locate = ($location == "") ? "" : "AND LOCATION = '".$location. "' ";
    
    $stmt = $conn->prepare('SELECT PRODUCT,SUM(ENDING_PALLET) AS ENDING_PALLET,SUM(ENDING_CASE) AS ENDING_CASE,SUM(ENDING_BOTTLE) AS ENDING_BOTTLE,SUM(ENDING_PARTIAL) AS ENDING_PARTIAL FROM `total` WHERE LOCATION != :loc AND `date` BETWEEN :dt1 and :dt2 AND LOCATION IN (SELECT NAME FROM location WHERE CHARGEABLE = :chrg) '.$locate.'GROUP BY PRODUCT');
    $stmt->execute(array('loc'=>'Boneyard','dt1'=>$date['dateFrom'],'dt2'=>$date['dateTo'],'chrg'=>'Chargeable'));
    
    $total_client_data = array();
    
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {        
        $total_client_data['ENDING_PALLET'][strtolower($row['PRODUCT'])] = $row['ENDING_PALLET'];
        $total_client_data['ENDING_CASE'][strtolower($row['PRODUCT'])] = $row['ENDING_CASE'];
        $total_client_data['ENDING_BOTTLE'][strtolower($row['PRODUCT'])] = $row['ENDING_BOTTLE'];
        $total_client_data['ENDING_PARTIAL'][strtolower($row['PRODUCT'])] = $row['ENDING_PARTIAL'];
    }
    return $total_client_data;
}

function getTotalSales( $date,$time,$location = "" ) {
    global $conn;
    $locate = ($location == "") ? "" : "AND LOCATION = '".$location. "'";
    $stmt = $conn->prepare('SELECT SUM(CASH) AS CASH,SUM(OTHER) AS OTHER,SUM(CREDIT) AS CREDIT FROM `total` WHERE `date` BETWEEN :dt1 and :dt2 AND LOCATION != :loc '.$locate);
    $stmt->execute(array('loc'=>'Boneyard','dt1'=>$date['dateFrom'],'dt2'=>$date['dateTo']));
    $total_sales_data = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $total_sales_data['CASH'] = $row['CASH'];
        $total_sales_data['OTHER'] = $row['OTHER'];
        $total_sales_data['CREDIT'] = $row['CREDIT'];
    }
    return $total_sales_data;
}

function getBegFieldInv()
{    
    global $conn;
    $stmt = $conn->prepare('SELECT PRODUCT,SUM(`qty_pallet`) as QTY_PALLET,SUM(`qty_case`) as QTY_CASE,SUM(`qty_individual`) as QTY_INDIV,SUM(`qty_partial`) as QTY_PARTIAL FROM `locaproduct` WHERE location in (SELECT NAME FROM location WHERE NOTES=:notes) AND LOCATION != :loc GROUP BY PRODUCT');
    $stmt->execute(array('loc'=>'Boneyard','notes'=>'Field'));
    $beg_field_inv = array();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $beg_field_inv['PRODUCT'][] = $row['PRODUCT'];
        $beg_field_inv[$row['PRODUCT']]['QTY_PALLET'] = $row['QTY_PALLET'];
        $beg_field_inv[$row['PRODUCT']]['QTY_CASE'] = $row['QTY_CASE'];
        $beg_field_inv[$row['PRODUCT']]['QTY_INDIV'] = $row['QTY_INDIV'];
        $beg_field_inv[$row['PRODUCT']]['QTY_PARTIAL'] = $row['QTY_PARTIAL'];
    }
    return $beg_field_inv;
}

function getPAndLInfo($product_ids,$filtered_product_ids,$transfers,$total_data,$total_spoilage_data,$total_client_data,$total_sales_data,$beg_field_inv,$total_received_data,$subcomm_data,$prior_total_data,$prior_subcomm_data,$prior_client_data) {
    global $conn, $objPHPExcel,$date; 
    $rowCount = 5;
    $stmt = $conn->prepare('SELECT ID,NAME,SIZE,QTY_PER_CASE,QTY_PER_PALLET,OPENING_PALLET,OPENING_CASE,OPENING_BOTTLE,OPENING_PARTIAL,CUR_PALLET,CUR_CASE,CUR_BOTTLE,CUR_PARTIAL,DRINKS_INDIVIDUALS,SALEPRICE FROM `product` WHERE NOTES != :notes ');
    $stmt->execute(array('notes'=>'Kegs'));

    /** Writing all the heading values * */
    
    $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Total Venue Control LLC.');
    $objPHPExcel->getActiveSheet()->SetCellValue('A2', 'P and L Report');
    $objPHPExcel->getActiveSheet()->SetCellValue('A3', "Date: {$date['dateFrom']} thru {$date['dateTo']}");    
    
    $objPHPExcel->getActiveSheet()->SetCellValue('H3', 'Interstand Transfers');    
    
    $objPHPExcel->getActiveSheet()->SetCellValue('A4', 'Product Description');
    $objPHPExcel->getActiveSheet()->SetCellValue('B4', '(+) Boneyard Daily Inventory Receiving');
    //$objPHPExcel->getActiveSheet()->SetCellValue('C4', 'Plus: Field Beginning Inventories');
    $objPHPExcel->getActiveSheet()->SetCellValue('C4', '(+) Prior Day Boneyard Ending Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('D4', '(+) Prior Day Sub Comm. Ending Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('E4', '(+) Prior Day Client Field Ending Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('F4', '(+)"RNC" Returns from Non-Charageable to Boneyard');
    $objPHPExcel->getActiveSheet()->SetCellValue('G4', '(-)"NC" Boneyard Transfers to Non-Charageable');
    $objPHPExcel->getActiveSheet()->SetCellValue('H4', 'Tranfers In (+) "IST" NC to C Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('I4', 'Tranfers Out (-) "IST" C to NC Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('J4', '(-)Client Boneyard Ending Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('K4', '(-) Client Sub Comm. Ending Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('L4', '(-)"C" Client Field Ending Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('M4', '(-)Spoilage  (Boneyard and Field)');
    $objPHPExcel->getActiveSheet()->SetCellValue('N4', '(=)Client Exp. Sales in Units');
    $objPHPExcel->getActiveSheet()->SetCellValue('O4', 'Units of Sale Per Item');
    $objPHPExcel->getActiveSheet()->SetCellValue('P4', 'Price Per Unit of Sale');
    $objPHPExcel->getActiveSheet()->SetCellValue('Q4', '(=)Cal. Rev. Generated Per Sales Item');
    
    $styleArray = array(
      'borders' => array(
          'allborders' => array(
              'style' => PHPExcel_Style_Border::BORDER_MEDIUM
          )
      )
  );
//    $p_total_data = array();
//    if($date['dateFrom'] == $date['dateTo']) {
//        $date1['dateFrom'] = date('Y-m-d', strtotime($date['dateFrom'] .' -1 day'));
//        $date1['dateTo'] = date('Y-m-d', strtotime($date['dateFrom'] .' -1 day'));
//        //echo "date : ";
//        //var_dump($date1);
//        $p_total_data = getTotalData(array(),$date1);
//        //echo "if";
//    }
    //var_dump($p_total_data);
      //  exit();
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        if($date['dateFrom'] == $date['dateTo']) {
            $total_received = ($row['QTY_PER_PALLET'] * $total_received_data[$row['NAME']]['total_pallet'] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $total_received_data[$row['NAME']]['total_case']) + $total_received_data[$row['NAME']]['total_individual'] + $total_received_data[$row['NAME']]['total_partial'];
        } else {
            $total_received = ($row['QTY_PER_PALLET'] * $row['OPENING_PALLET'] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $row['OPENING_CASE']) + $row['OPENING_BOTTLE'] + $row['OPENING_PARTIAL'];
        }
       // $total_received = ($row['QTY_PER_PALLET'] * $total_received_data[$row['NAME']]['total_pallet'] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $total_received_data[$row['NAME']]['total_case']) + $total_received_data[$row['NAME']]['total_individual'] + $total_received_data[$row['NAME']]['total_partial'];
        $to_nc = '-';        
        $from_nc = '-';
        
//        if(!empty($p_total_data)) {
//            $p_aei = (($row['QTY_PER_PALLET'] * $p_total_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $p_total_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($p_total_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $p_total_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
//            $total_received = $total_received + $p_aei;
//        }
        
        if (in_array($row['ID'], $product_ids)) {
           
            if (in_array($row['ID'], $filtered_product_ids['to_nc']['PRODUCT_ID'])) {
                $to_nc = (($row['QTY_PER_PALLET'] * $filtered_product_ids['to_nc']['QTY_PALLET'][$row['ID']] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $filtered_product_ids['to_nc']['QTY_CASE'][$row['ID']]) + ($filtered_product_ids['to_nc']['QTY_BOTTLE'][$row['ID']] + $filtered_product_ids['to_nc']['QTY_PARTIAL'][$row['ID']]));
            }
       
            if (in_array($row['ID'], $filtered_product_ids['from_nc']['PRODUCT_ID'])) {
                $from_nc = (($row['QTY_PER_PALLET'] * $filtered_product_ids['from_nc']['QTY_PALLET'][$row['ID']] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $filtered_product_ids['from_nc']['QTY_CASE'][$row['ID']]) + ($filtered_product_ids['from_nc']['QTY_BOTTLE'][$row['ID']] + $filtered_product_ids['from_nc']['QTY_PARTIAL'][$row['ID']]));
            }
           
        }
        
        $total_beg_field_inv = 0;
        if(in_array($row['NAME'], $beg_field_inv['PRODUCT']))
        {
            $total_beg_field_inv = ($row['QTY_PER_PALLET'] * $beg_field_inv[$row['NAME']]['QTY_PALLET'] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $beg_field_inv[$row['NAME']]['QTY_CASE']) + $beg_field_inv[$row['NAME']]['QTY_INDIV'] + $beg_field_inv[$row['NAME']]['QTY_PARTIAL'];
        }
        
        $prior_aei = (($row['QTY_PER_PALLET'] * $prior_total_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $prior_total_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($prior_total_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $prior_total_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
        $aei = (($row['QTY_PER_PALLET'] * $total_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $total_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($total_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $total_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
        $subcomm_aei = (($row['QTY_PER_PALLET'] * $subcomm_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $subcomm_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($subcomm_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $subcomm_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
        $prior_subcomm = (($row['QTY_PER_PALLET'] * $prior_subcomm_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $prior_subcomm_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($prior_subcomm_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $prior_subcomm_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
        $cfe = (($row['QTY_PER_PALLET'] * $total_client_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $total_client_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($total_client_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $total_client_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
        $prior_cfe = (($row['QTY_PER_PALLET'] * $prior_client_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $prior_client_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($prior_client_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $prior_client_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
        
        $objPHPExcel->getActiveSheet()->getStyle('A'.$rowCount.':Q'.$rowCount)->getNumberFormat()->setFormatCode(
            '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        );
        
        $objPHPExcel->getActiveSheet()->getStyle('J'.$rowCount.':M'.$rowCount)->getNumberFormat()->setFormatCode(
            '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        );
        
        $objPHPExcel->getActiveSheet()->getStyle('C'.$rowCount.':I'.$rowCount)->getNumberFormat()->setFormatCode(
            '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        );
        
        $objPHPExcel->getActiveSheet()->getStyle('N'.$rowCount.':O'.$rowCount)->getNumberFormat()->setFormatCode(
            '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        );
        
        $objPHPExcel->getActiveSheet()->getStyle('P'.$rowCount.':Q'.$rowCount)->getNumberFormat()->setFormatCode(
            '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        );
              
        $h_val = $transfers['NCToC'][$row['ID']]['qty_case'] * $row['QTY_PER_CASE'] + $transfers['NCToC'][$row['ID']]['qty_pallet'] * $row['QTY_PER_CASE'] * $row['QTY_PER_PALLET'] + $transfers['NCToC'][$row['ID']]['qty_bottle'] + $transfers['NCToC'][$row['ID']]['qty_partial'];        
        $i_val = $transfers['CToNC'][$row['ID']]['qty_case'] * $row['QTY_PER_CASE'] + $transfers['CToNC'][$row['ID']]['qty_pallet'] * $row['QTY_PER_CASE'] * $row['QTY_PER_PALLET'] + $transfers['CToNC'][$row['ID']]['qty_bottle'] + $transfers['CToNC'][$row['ID']]['qty_partial'];
        
        
        $objPHPExcel->getActiveSheet()->getStyle("A$rowCount:R$rowCount")->getFont()->setBold(true);
        
        /** Setting all the values in the report * */
        $objPHPExcel->getActiveSheet()->SetCellValue('A' . $rowCount, $row['NAME']);
        $objPHPExcel->getActiveSheet()->SetCellValue('B' . $rowCount, $total_received);
        //$objPHPExcel->getActiveSheet()->SetCellValue('C' . $rowCount, $total_beg_field_inv);
        $objPHPExcel->getActiveSheet()->SetCellValue('C' . $rowCount, $prior_aei);
        $objPHPExcel->getActiveSheet()->SetCellValue('D' . $rowCount, $prior_subcomm);
        $objPHPExcel->getActiveSheet()->SetCellValue('E' . $rowCount, $prior_cfe-$prior_subcomm);
        $objPHPExcel->getActiveSheet()->SetCellValue('F' . $rowCount, $from_nc);
        $objPHPExcel->getActiveSheet()->SetCellValue('G' . $rowCount, -$to_nc);
        $objPHPExcel->getActiveSheet()->SetCellValue('H' . $rowCount, $h_val);
        $objPHPExcel->getActiveSheet()->SetCellValue('I' . $rowCount, -$i_val);
        $objPHPExcel->getActiveSheet()->SetCellValue('J' . $rowCount, -$aei);
        $objPHPExcel->getActiveSheet()->SetCellValue('K' . $rowCount, -$subcomm_aei);
        $objPHPExcel->getActiveSheet()->SetCellValue('L' . $rowCount, -$cfe+$subcomm_aei);
        $objPHPExcel->getActiveSheet()->SetCellValue('M' . $rowCount, -$total_spoilage_data[$row['NAME']]);
        $objPHPExcel->getActiveSheet()->SetCellValue('N' . $rowCount, '=SUM(B'.$rowCount.':M'.$rowCount.')');
        $objPHPExcel->getActiveSheet()->SetCellValue('O' . $rowCount, $row['DRINKS_INDIVIDUALS']);
        $objPHPExcel->getActiveSheet()->SetCellValue('P' . $rowCount, $row['SALEPRICE']);
        $objPHPExcel->getActiveSheet()->SetCellValue('Q' . $rowCount, '=N'.$rowCount.'*O'.$rowCount.'*P'.$rowCount);
        
        $objPHPExcel->getActiveSheet()->getStyle("Q".$rowCount)->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '00FF00')
                )
            )
        );
        
        $objPHPExcel->getActiveSheet()->getStyle('A'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $objPHPExcel->getActiveSheet()->getStyle('B'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('C'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('D'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('E'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('F'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('G'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('H'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('I'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('J'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('K'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('L'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('M'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('N'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('O'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('P'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('Q'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        //$objPHPExcel->getActiveSheet()->getStyle('R'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        
        $objPHPExcel->getActiveSheet()->getStyle("A$rowCount:Q$rowCount")->applyFromArray($styleArray);
        
        $rowCount++;
    }
    $rowCount--;
    $objPHPExcel->getActiveSheet()->SetCellValue("Q".($rowCount+1), '=SUM(Q3:Q'.($rowCount).')' );
    $objPHPExcel->getActiveSheet()->SetCellValue("Q".($rowCount+2), $total_sales_data['CASH'] );
    $objPHPExcel->getActiveSheet()->SetCellValue("Q".($rowCount+3), $total_sales_data['CREDIT'] );
    $objPHPExcel->getActiveSheet()->SetCellValue("Q".($rowCount+4), $total_sales_data['OTHER'] );
    $objPHPExcel->getActiveSheet()->SetCellValue("Q".($rowCount+5), '=SUM(Q'.($rowCount+2).':Q'.($rowCount+4).')');
    $objPHPExcel->getActiveSheet()->SetCellValue("Q".($rowCount+6), '=Q'.($rowCount+1).'-Q'.($rowCount+5));
    $objPHPExcel->getActiveSheet()->SetCellValue("Q".($rowCount+7), '=Q'.($rowCount+6).'/Q'.($rowCount+5));



//    $objPHPExcel->getActiveSheet()->getStyle( 'Q'.($rowCount+7) )->getNumberFormat()->setFormatCode(
//            '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
//        );


    $objPHPExcel->getActiveSheet()->getStyle( 'Q'.($rowCount+7) )
    ->getNumberFormat()->applyFromArray( 
        array( 
            'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00
        )
    );

    $objPHPExcel->getActiveSheet()->getStyle( 'Q'.($rowCount+7) )->getNumberFormat()->setFormatCode(
        '_(0%_);(0%);_(0%_);@'
    );

    $objPHPExcel->getActiveSheet()->getStyle("Q".($rowCount+7))->getFont()->setSize(12)->getColor()->setRGB('FF0000');;
    
    
    $objPHPExcel->getActiveSheet()->getStyle("Q".($rowCount+1).":Q".($rowCount+7))->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("Q".($rowCount+1).":Q".($rowCount+6))->getNumberFormat()->setFormatCode(
            '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        );
    
   $styleArray2 = array(
	'borders' => array(
		'allborders' => array(
			'style' => PHPExcel_Style_Border::BORDER_MEDIUM			
		),
	),
    );
    $objPHPExcel->getActiveSheet()->getStyle("Q".($rowCount+1).":Q".($rowCount+7))->applyFromArray($styleArray2);
    
    $objPHPExcel->getActiveSheet()->getStyle("Q".($rowCount+1))->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => '00FF00')
            )
        )
    );
    
     $objPHPExcel->getActiveSheet()->getStyle("Q".($rowCount+5))->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => '00FF00')
            )
        )
    );
     
     $objPHPExcel->getActiveSheet()->getStyle("Q".($rowCount+6).":Q".($rowCount+7))->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => 'FFFF00')
            )
        )
    );
    
    $start = ++$rowCount;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Total "Sales Based Upon Inventory": ');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => '00FF00')
            )
        )
    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Total Cash Sales (less starting banks): ');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
//    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
//    array(
//        'fill' => array(
//            'type' => PHPExcel_Style_Fill::FILL_SOLID,
//            'color' => array('rgb' => 'FDFDD9')
//            )
//        )
//    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Total Credit Card Sales: ');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
//    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
//    array(
//        'fill' => array(
//            'type' => PHPExcel_Style_Fill::FILL_SOLID,
//            'color' => array('rgb' => 'FDFDD9')
//            )
//        )
//    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Total Other Sales (e.g. comp tickets, etc.): ');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
//    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
//    array(
//        'fill' => array(
//            'type' => PHPExcel_Style_Fill::FILL_SOLID,
//            'color' => array('rgb' => 'FDFDD9')
//            )
//        )
//    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Total "Sales Based Upon Payments": ');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => '00FF00')
            )
        )
    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Over (Short) in $s');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => 'F2F2F2')
            )
        )
    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Over (Short) in %');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => 'F2F2F2')
            )
        )
    );
    
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A$start:P$start");
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+1).":P".($start+1));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+2).":P".($start+2));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+3).":P".($start+3));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+4).":P".($start+4));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+5).":P".($start+5));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+6).":P".($start+6));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+7).":P".($start+7));



     $styleArray1 = array(
	'borders' => array(
		'allborders' => array(
			'style' => PHPExcel_Style_Border::BORDER_MEDIUM			
		),
	),
    );
    $objPHPExcel->getActiveSheet()->getStyle("A$start:"."P".($start))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+1).":P".($start+1))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+2).":P".($start+2))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+3).":P".($start+3))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+4).":P".($start+4))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+5).":P".($start+5))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+6).":P".($start+6))->applyFromArray($styleArray1);
    
    /** Formatting all the heading columns * */
    $objPHPExcel->getActiveSheet()->getStyle("H3:I3")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("A4:Q4")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("A1:Q3")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(32.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(10.4);
    $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(10.71);
    $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(10.29);
    $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(10.71);
    $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(10.50);
    $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(10.50);
    $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth(15.00);
    //$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth(15.43);
    
    $objPHPExcel->getActiveSheet()->getRowDimension('4')->setRowHeight(84.75);
    $objPHPExcel->getActiveSheet()->getStyle('B4:Q4')
    ->getAlignment()->setWrapText(true); 
    $objPHPExcel->getActiveSheet()->getStyle("B4:Q4")->getFont()->setSize(10);
    $objPHPExcel->getActiveSheet()->getStyle('A4:Q4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle('A4')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle('H3:I3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle('A4:Q4')->applyFromArray($styleArray);
    
    $objPHPExcel->getActiveSheet()->getStyle('A4:P4')->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'F2F2F2')
            )
        )
    );
    $objPHPExcel->getActiveSheet()->getStyle('Q4')->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '00FF00')
            )
        )
    );
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("H3:I3");
    
//    $objPHPExcel->getActiveSheet()->getStyle('H3:I3')->applyFromArray(
//        array(
//            'fill' => array(
//                'type' => PHPExcel_Style_Fill::FILL_SOLID,
//                'color' => array('rgb' => 'FDFDD9')
//            )
//        )
//    );
        
    
    
    $styleArray3 = array(
	'borders' => array(
		'outline' => array(
			'style' => PHPExcel_Style_Border::BORDER_MEDIUM			
		),
	),
    );
    $objPHPExcel->getActiveSheet()->getStyle("H3:I3")->applyFromArray($styleArray3);
    
    /** Saving the file * */
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    ob_end_clean();
    $objWriter->save('php://output');
}


function standBegFieldInv($loc,$date){
    global $conn;
    $current_date = date('Y-m-d');
     
//    if(in_array($current_date,$date)) {
//        $stmt = $conn->prepare("SELECT `product`,`total_pallet`,`total_case`,`total_individual`,`total_partial` FROM `locaproduct` WHERE `location` = :loc");
//        $stmt->execute(array('loc'=>$loc));
//        $curstandBeg = array();
//        while($row = $stmt->fetch(PDO::FETCH_ASSOC))
//        {
//            $curstandBeg[$row['product']]['qty_pallet'] = $row['total_pallet'];
//            $curstandBeg[$row['product']]['qty_case'] = $row['total_case'];
//            $curstandBeg[$row['product']]['qty_individual'] = $row['total_individual'];
//            $curstandBeg[$row['product']]['qty_partial'] = $row['total_partial'];
//        }        
//    }
//    $stmt = $conn->prepare("SELECT `product`,SUM(`total_pallet`) as `total_pallet` ,SUM(`total_case`) as `total_case`,SUM(`total_individual`) as `total_individual`,SUM(`total_partial`) as `total_partial` FROM `totaldate` WHERE `location` = :loc and `date` BETWEEN :dt1 and :dt2 GROUP BY `product`");
//    $stmt->execute(array('loc'=>$loc,'dt1'=>$date['dateFrom'],'dt2'=>$date['dateTo']));
    
    $stmt = $conn->prepare("SELECT `product_id`,SUM(`qty_pallet`) as `total_pallet` ,SUM(`qty_case`) as `total_case`,SUM(`qty_bottle`) as `total_individual`,SUM(`qty_partial`) as `total_partial` FROM `bill` WHERE `order_id` IN (SELECT `order_id` FROM `order` WHERE `location_from` = 'Boneyard' and `location_to` = :loc AND DATE(`date`) BETWEEN :dt1 and :dt2) GROUP BY `product_id`");
    $stmt->execute(array('loc'=>$loc,'dt1'=>$date['dateFrom'],'dt2'=>$date['dateTo']));
    
    $standBeg = array();
    while($row = $stmt->fetch(PDO::FETCH_ASSOC))
    {
//        if(empty($curstandBeg)) {
            $standBeg[$row['product_id']]['qty_pallet'] = $row['total_pallet'];
            $standBeg[$row['product_id']]['qty_case'] = $row['total_case'];
            $standBeg[$row['product_id']]['qty_individual'] = $row['total_individual'];
            $standBeg[$row['product_id']]['qty_partial'] = $row['total_partial'];
//        } else {
//            $standBeg[$row['product']]['qty_pallet'] = $row['total_pallet'] + $curstandBeg[$row['product']]['qty_pallet'];
//            $standBeg[$row['product']]['qty_case'] = $row['total_case'] + $curstandBeg[$row['product']]['qty_case'];
//            $standBeg[$row['product']]['qty_individual'] = $row['total_individual'] + $curstandBeg[$row['product']]['qty_individual'];
//            $standBeg[$row['product']]['qty_partial'] = $row['total_partial'] + $curstandBeg[$row['product']]['qty_partial'];
//        }
    }
    
//    if(empty($standBeg)) {
//        $standBeg = $curstandBeg;
//    }
    
    return $standBeg;
}


function getPreviousDayValue($product_ids,$filtered_product_ids,$transfers,$total_data,$total_spoilage_data,$total_client_data,$total_sales_data,$beg_field_inv,$standBeg) {
    global $conn;
    $stmt = $conn->prepare('SELECT ID,NAME,SIZE,QTY_PER_CASE,QTY_PER_PALLET,OPENING_PALLET,OPENING_CASE,OPENING_BOTTLE,OPENING_PARTIAL,CUR_PALLET,CUR_CASE,CUR_BOTTLE,CUR_PARTIAL,DRINKS_INDIVIDUALS,SALEPRICE FROM `product` WHERE NOTES != :notes ');
    $stmt->execute(array('notes'=>'Kegs'));

    
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $total_received = ($row['QTY_PER_PALLET'] * $standBeg[$row['NAME']]['qty_pallet'] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $standBeg[$row['NAME']]['qty_case']) + $standBeg[$row['NAME']]['qty_individual'] + $standBeg[$row['NAME']]['qty_partial'];        
        if (in_array($row['ID'], $product_ids)) {
           
            if (in_array($row['ID'], $filtered_product_ids['to_nc']['PRODUCT_ID'])) {
                $to_nc = (($row['QTY_PER_PALLET'] * $filtered_product_ids['to_nc']['QTY_PALLET'][$row['ID']] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $filtered_product_ids['to_nc']['QTY_CASE'][$row['ID']]) + ($filtered_product_ids['to_nc']['QTY_BOTTLE'][$row['ID']] + $filtered_product_ids['to_nc']['QTY_PARTIAL'][$row['ID']]));
            }
       
            if (in_array($row['ID'], $filtered_product_ids['from_nc']['PRODUCT_ID'])) {
                $from_nc = (($row['QTY_PER_PALLET'] * $filtered_product_ids['from_nc']['QTY_PALLET'][$row['ID']] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $filtered_product_ids['from_nc']['QTY_CASE'][$row['ID']]) + ($filtered_product_ids['from_nc']['QTY_BOTTLE'][$row['ID']] + $filtered_product_ids['from_nc']['QTY_PARTIAL'][$row['ID']]));
            }
           
        }
        
        $total_beg_field_inv = 0;
        if(in_array($row['NAME'], $beg_field_inv['PRODUCT']))
        {
            $total_beg_field_inv = ($row['QTY_PER_PALLET'] * $beg_field_inv[$row['NAME']]['QTY_PALLET'] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $beg_field_inv[$row['NAME']]['QTY_CASE']) + $beg_field_inv[$row['NAME']]['QTY_INDIV'] + $beg_field_inv[$row['NAME']]['QTY_PARTIAL'];
        }
        
        
        $aei = (($row['QTY_PER_PALLET'] * $total_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $total_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($total_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $total_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
        $cfe = (($row['QTY_PER_PALLET'] * $total_client_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $total_client_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($total_client_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $total_client_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
              
        $value[$row['NAME']] = $total_received -( $transfers['CToNC'][$row['ID']]['qty_case'] * $row['QTY_PER_CASE'] + $transfers['CToNC'][$row['ID']]['qty_pallet'] * $row['QTY_PER_CASE'] * $row['QTY_PER_PALLET'] + $transfers['CToNC'][$row['ID']]['qty_bottle'] + $transfers['CToNC'][$row['ID']]['qty_partial'] ) -( $transfers['CToC'][$row['ID']]['qty_case'] * $row['QTY_PER_CASE'] + $transfers['CToC'][$row['ID']]['qty_pallet'] * $row['QTY_PER_CASE'] * $row['QTY_PER_PALLET'] + $transfers['CToC'][$row['ID']]['qty_bottle'] + $transfers['CToC'][$row['ID']]['qty_partial'] ) -$cfe-$total_spoilage_data[$row['NAME']];
        
    }
    return $value;
}


function getStandPAndLInfo($product_ids,$filtered_product_ids,$transfers,$total_data,$total_spoilage_data,$total_client_data,$total_sales_data,$beg_field_inv,$standBeg,$val,$prior_cfe_data) {
    global $conn, $objPHPExcel, $rowCount, $location, $date;
    $stmt = $conn->prepare('SELECT ID,NAME,SIZE,QTY_PER_CASE,QTY_PER_PALLET,OPENING_PALLET,OPENING_CASE,OPENING_BOTTLE,OPENING_PARTIAL,CUR_PALLET,CUR_CASE,CUR_BOTTLE,CUR_PARTIAL,DRINKS_INDIVIDUALS,SALEPRICE FROM `product` WHERE NOTES != :notes ');
    $stmt->execute(array('notes'=>'Kegs'));

    /** Writing all the heading values * */
    $objPHPExcel->getActiveSheet()->SetCellValue('A1', "Total Venue Control LLC.");
    $objPHPExcel->getActiveSheet()->SetCellValue('A2', "Stand P & L Report" );
    $objPHPExcel->getActiveSheet()->SetCellValue('A3', "Stand Name: ".$location );    
    $objPHPExcel->getActiveSheet()->SetCellValue('A4', "Date: {$date['dateFrom']} thru {$date['dateTo']}");    
    
    $objPHPExcel->getActiveSheet()->SetCellValue('D4', 'Interstand Transfers');    
    
    $objPHPExcel->getActiveSheet()->SetCellValue('A5', 'Product Description');
    $objPHPExcel->getActiveSheet()->SetCellValue('B5', '(+) Stand Daily Inventory Receiving');
    $objPHPExcel->getActiveSheet()->SetCellValue('C5', '(+) Prior Day Stand Ending Inventory');
    //$objPHPExcel->getActiveSheet()->SetCellValue('C2', 'Plus: Field Beginning Inventories');
    //$objPHPExcel->getActiveSheet()->SetCellValue('D2', '(+)"RNC" Returns from Non-Client to Stand');
    //$objPHPExcel->getActiveSheet()->SetCellValue('E2', '(-)"NC" Transfers to Non-Client to Stand');
    $objPHPExcel->getActiveSheet()->SetCellValue('D5', '(+) Transfers IN "IST" NC to C Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('E5', '(+) Transfers IN "IST" C to C Inventory');

    $objPHPExcel->getActiveSheet()->SetCellValue('F5', '(-)Transfers OUT "IST" C to NC Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('G5', '(-)Transfers OUT "IST" C to C Inventory');
    //$objPHPExcel->getActiveSheet()->SetCellValue('H2', '(-)Client Boneyard Ending Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('H5', '(-)Client Stand Ending Inventory');
    $objPHPExcel->getActiveSheet()->SetCellValue('I5', '(-)Spoilage');
    $objPHPExcel->getActiveSheet()->SetCellValue('J5', '(=)Client Exp. Sales in Units');
    $objPHPExcel->getActiveSheet()->SetCellValue('K5', 'Units of Sale Per Item');
    $objPHPExcel->getActiveSheet()->SetCellValue('L5', 'Price Per Unit of Sale');
    $objPHPExcel->getActiveSheet()->SetCellValue('M5', '(=)Cal. Rev. Generated Per Sales Item');
    
    $styleArray = array(
      'borders' => array(
          'allborders' => array(
              'style' => PHPExcel_Style_Border::BORDER_MEDIUM
          )
      )
  );
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        $total_received = ($row['QTY_PER_PALLET'] * $standBeg[$row['ID']]['qty_pallet'] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $standBeg[$row['ID']]['qty_case']) + $standBeg[$row['ID']]['qty_individual'] + $standBeg[$row['ID']]['qty_partial'];
        //$total_received = $total_received + $val[$row['NAME']];
        $to_nc = '-';        
        $from_nc = '-';

        if (in_array($row['ID'], $product_ids)) {
           
            if (in_array($row['ID'], $filtered_product_ids['to_nc']['PRODUCT_ID'])) {
                $to_nc = (($row['QTY_PER_PALLET'] * $filtered_product_ids['to_nc']['QTY_PALLET'][$row['ID']] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $filtered_product_ids['to_nc']['QTY_CASE'][$row['ID']]) + ($filtered_product_ids['to_nc']['QTY_BOTTLE'][$row['ID']] + $filtered_product_ids['to_nc']['QTY_PARTIAL'][$row['ID']]));
            }
       
            if (in_array($row['ID'], $filtered_product_ids['from_nc']['PRODUCT_ID'])) {
                $from_nc = (($row['QTY_PER_PALLET'] * $filtered_product_ids['from_nc']['QTY_PALLET'][$row['ID']] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $filtered_product_ids['from_nc']['QTY_CASE'][$row['ID']]) + ($filtered_product_ids['from_nc']['QTY_BOTTLE'][$row['ID']] + $filtered_product_ids['from_nc']['QTY_PARTIAL'][$row['ID']]));
            }
           
        }
        
        $total_beg_field_inv = 0;
        if(in_array($row['NAME'], $beg_field_inv['PRODUCT']))
        {
            $total_beg_field_inv = ($row['QTY_PER_PALLET'] * $beg_field_inv[$row['NAME']]['QTY_PALLET'] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $beg_field_inv[$row['NAME']]['QTY_CASE']) + $beg_field_inv[$row['NAME']]['QTY_INDIV'] + $beg_field_inv[$row['NAME']]['QTY_PARTIAL'];
        }
        
        
        $aei = (($row['QTY_PER_PALLET'] * $total_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $total_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($total_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $total_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
        
        $prior_cfe = (($row['QTY_PER_PALLET'] * $prior_cfe_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $prior_cfe_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($prior_cfe_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $prior_cfe_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
        $cfe = (($row['QTY_PER_PALLET'] * $total_client_data['ENDING_PALLET'][strtolower($row['NAME'])] * $row['QTY_PER_CASE']) + ($row['QTY_PER_CASE'] * $total_client_data['ENDING_CASE'][strtolower($row['NAME'])]) + ($total_client_data['ENDING_BOTTLE'][strtolower($row['NAME'])] + $total_client_data['ENDING_PARTIAL'][strtolower($row['NAME'])]));
        
        $objPHPExcel->getActiveSheet()->getStyle('A'.$rowCount.':N'.$rowCount)->getNumberFormat()->setFormatCode(
            '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        );
        
        $objPHPExcel->getActiveSheet()->getStyle('F'.$rowCount.':M'.$rowCount)->getNumberFormat()->setFormatCode(
            '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        );
        
        $objPHPExcel->getActiveSheet()->getStyle('B'.$rowCount.':G'.$rowCount)->getNumberFormat()->setFormatCode(
            '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        );
        
        $objPHPExcel->getActiveSheet()->getStyle('L'.$rowCount.':M'.$rowCount)->getNumberFormat()->setFormatCode(
            '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
            //'_("$"* ??_);_(@_)'
        );
        
        $d_val = $transfers['NCToC'][$row['ID']]['qty_case'] * $row['QTY_PER_CASE'] + $transfers['NCToC'][$row['ID']]['qty_pallet'] * $row['QTY_PER_CASE'] * $row['QTY_PER_PALLET'] + $transfers['NCToC'][$row['ID']]['qty_bottle'] + $transfers['NCToC'][$row['ID']]['qty_partial'];
        $e_val = $transfers['PCToC'][$row['ID']]['qty_case'] * $row['QTY_PER_CASE'] + $transfers['PCToC'][$row['ID']]['qty_pallet'] * $row['QTY_PER_CASE'] * $row['QTY_PER_PALLET'] + $transfers['PCToC'][$row['ID']]['qty_bottle'] + $transfers['PCToC'][$row['ID']]['qty_partial'];
        $f_val = $transfers['CToNC'][$row['ID']]['qty_case'] * $row['QTY_PER_CASE'] + $transfers['CToNC'][$row['ID']]['qty_pallet'] * $row['QTY_PER_CASE'] * $row['QTY_PER_PALLET'] + $transfers['CToNC'][$row['ID']]['qty_bottle'] + $transfers['CToNC'][$row['ID']]['qty_partial'];
        $g_val = $transfers['CToC'][$row['ID']]['qty_case'] * $row['QTY_PER_CASE'] + $transfers['CToC'][$row['ID']]['qty_pallet'] * $row['QTY_PER_CASE'] * $row['QTY_PER_PALLET'] + $transfers['CToC'][$row['ID']]['qty_bottle'] + $transfers['CToC'][$row['ID']]['qty_partial'];

        if( $prior_cfe == 0 && $d_val == 0 && $e_val == 0 && $f_val == 0 && $g_val == 0 && $cfe == 0 && $total_spoilage_data[$row['NAME']] == 0 ) {
            $objPHPExcel->getActiveSheet()->getRowDimension($rowCount)->setVisible(false);
        }

        $objPHPExcel->getActiveSheet()->getStyle("A$rowCount:M$rowCount")->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle("A1:A4")->getFont()->setBold(true);
                
        /** Setting all the values in the report * */
        $objPHPExcel->getActiveSheet()->SetCellValue('A' . $rowCount, $row['NAME']);
        $objPHPExcel->getActiveSheet()->SetCellValue('B' . $rowCount, $total_received);
        $objPHPExcel->getActiveSheet()->SetCellValue('C' . $rowCount, $prior_cfe);
       // $objPHPExcel->getActiveSheet()->SetCellValue('D' . $rowCount, $from_nc);
       // $objPHPExcel->getActiveSheet()->SetCellValue('E' . $rowCount, -$to_nc);
        $objPHPExcel->getActiveSheet()->SetCellValue('D' . $rowCount, $d_val);
        $objPHPExcel->getActiveSheet()->SetCellValue('E' . $rowCount, $e_val);
        $objPHPExcel->getActiveSheet()->SetCellValue('F' . $rowCount, -$f_val);
        $objPHPExcel->getActiveSheet()->SetCellValue('G' . $rowCount, -$g_val);
        //$objPHPExcel->getActiveSheet()->SetCellValue('H' . $rowCount, -$aei);
        $objPHPExcel->getActiveSheet()->SetCellValue('H' . $rowCount, -$cfe);
        $objPHPExcel->getActiveSheet()->SetCellValue('I' . $rowCount, -$total_spoilage_data[$row['NAME']]);
        $objPHPExcel->getActiveSheet()->SetCellValue('J' . $rowCount, '=SUM(A'.$rowCount.':I'.$rowCount.')');
        $objPHPExcel->getActiveSheet()->SetCellValue('K' . $rowCount, $row['DRINKS_INDIVIDUALS']);
        $objPHPExcel->getActiveSheet()->SetCellValue('L' . $rowCount, $row['SALEPRICE']);
        $objPHPExcel->getActiveSheet()->SetCellValue('M' . $rowCount, '=J'.$rowCount.'*K'.$rowCount.'*L'.$rowCount);
        
        $objPHPExcel->getActiveSheet()->getStyle("M".$rowCount)->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '00FF00')
                )
            )
        );
        
        $objPHPExcel->getActiveSheet()->getStyle('A'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $objPHPExcel->getActiveSheet()->getStyle('B'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('C'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('D'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('E'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('F'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('G'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('H'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('I'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('J'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('K'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('L'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $objPHPExcel->getActiveSheet()->getStyle('M'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        //$objPHPExcel->getActiveSheet()->getStyle('N'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        
        $objPHPExcel->getActiveSheet()->getStyle("A$rowCount:M$rowCount")->applyFromArray($styleArray);
        
        $rowCount++;
    }
    $rowCount--;
    $objPHPExcel->getActiveSheet()->SetCellValue("M".($rowCount+1), '=SUM(M3:M'.($rowCount).')' );
    $objPHPExcel->getActiveSheet()->SetCellValue("M".($rowCount+2), $total_sales_data['CASH'] );
    $objPHPExcel->getActiveSheet()->SetCellValue("M".($rowCount+3), $total_sales_data['CREDIT'] );
    $objPHPExcel->getActiveSheet()->SetCellValue("M".($rowCount+4), $total_sales_data['OTHER'] );
    $objPHPExcel->getActiveSheet()->SetCellValue("M".($rowCount+5), '=SUM(M'.($rowCount+2).':M'.($rowCount+4).')');
    $objPHPExcel->getActiveSheet()->SetCellValue("M".($rowCount+6), '=M'.($rowCount+1).'-M'.($rowCount+5));
    $objPHPExcel->getActiveSheet()->SetCellValue("M".($rowCount+7), '=M'.($rowCount+6).'/M'.($rowCount+5));
    
//    $objPHPExcel->getActiveSheet()->getStyle( 'M'.($rowCount+7) )->getNumberFormat()->setFormatCode(
//            '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
//        );
    $objPHPExcel->getActiveSheet()->getStyle( 'M'.($rowCount+7) )
    ->getNumberFormat()->applyFromArray( 
        array( 
            'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00
        )
    );

    $objPHPExcel->getActiveSheet()->getStyle( 'M'.($rowCount+7) )->getNumberFormat()->setFormatCode(
        '_(0%_);(0%);_(0%_);@'
    );

    $objPHPExcel->getActiveSheet()->getStyle("M".($rowCount+7))->getFont()->setSize(12)->getColor()->setRGB('FF0000');;
    
    
    $objPHPExcel->getActiveSheet()->getStyle("M".($rowCount+1).":M".($rowCount+7))->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("M".($rowCount+1).":M".($rowCount+6))->getNumberFormat()->setFormatCode(
            '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        );
    
   $styleArray2 = array(
	'borders' => array(
		'allborders' => array(
			'style' => PHPExcel_Style_Border::BORDER_MEDIUM			
		),
	),
    );
    $objPHPExcel->getActiveSheet()->getStyle("M".($rowCount+1).":M".($rowCount+7))->applyFromArray($styleArray2);
    
    $objPHPExcel->getActiveSheet()->getStyle("M".($rowCount+1))->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => '00FF00')
            )
        )
    );
    
    $objPHPExcel->getActiveSheet()->getStyle("M".($rowCount+5))->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => '00FF00')
            )
        )
    );
    
    $objPHPExcel->getActiveSheet()->getStyle("M".($rowCount+6).":M".($rowCount+7))->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => 'FFFF00')
            )
        )
    );
    
    $start = ++$rowCount;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Total "Sales Based Upon Inventory": ');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => '00FF00')
            )
        )
    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Total Cash Sales (less starting banks):	');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
//    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
//    array(
//        'fill' => array(
//            'type' => PHPExcel_Style_Fill::FILL_SOLID,
//            'color' => array('rgb' => 'FDFDD9')
//            )
//        )
//    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Total Credit Card Sales: ');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
//    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
//    array(
//        'fill' => array(
//            'type' => PHPExcel_Style_Fill::FILL_SOLID,
//            'color' => array('rgb' => 'FDFDD9')
//            )
//        )
//    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Total Other Sales (e.g. comp tickets, etc.): ');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
//    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
//    array(
//        'fill' => array(
//            'type' => PHPExcel_Style_Fill::FILL_SOLID,
//            'color' => array('rgb' => 'FDFDD9')
//            )
//        )
//    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Total "Sales Based Upon Payments": ');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => '00FF00')
            )
        )
    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Over (Short) in $s');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => 'F2F2F2')
            )
        )
    );
    $rowCount++;
    $objPHPExcel->getActiveSheet()->SetCellValue("A$rowCount", 'Over (Short) in %');
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("A$rowCount")->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => 'F2F2F2')
            )
        )
    );
    
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A$start:L$start");
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+1).":L".($start+1));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+2).":L".($start+2));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+3).":L".($start+3));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+4).":L".($start+4));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+5).":L".($start+5));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+6).":L".($start+6));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A".($start+7).":L".($start+7));
    
     $styleArray1 = array(
	'borders' => array(
		'allborders' => array(
			'style' => PHPExcel_Style_Border::BORDER_MEDIUM			
		),
	),
    );
    $objPHPExcel->getActiveSheet()->getStyle("A$start:"."M".($start))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+1).":M".($start+1))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+2).":M".($start+2))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+3).":M".($start+3))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+4).":M".($start+4))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+5).":M".($start+5))->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->getStyle("A".($start+6).":M".($start+6))->applyFromArray($styleArray1);
    
    $objPHPExcel->getActiveSheet()->getStyle("A5")->applyFromArray($styleArray1);
    
    /** Formatting all the heading columns * */
    $objPHPExcel->getActiveSheet()->getStyle("C4:D4")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getStyle("A5:N5")->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(40.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(12.30);
    $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(11.57);
    $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(12.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(11.97);
    $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(18.60);
    //$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(15.43);
    
    $objPHPExcel->getActiveSheet()->getRowDimension('5')->setRowHeight(69.75);
    $objPHPExcel->getActiveSheet()->getStyle('B5:M5')
    ->getAlignment()->setWrapText(true); 
    $objPHPExcel->getActiveSheet()->getStyle("B5:M5")->getFont()->setSize(10);
    $objPHPExcel->getActiveSheet()->getStyle('A5:M5')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle('A5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
     
    $objPHPExcel->getActiveSheet()->getStyle('A5:M5')->applyFromArray($styleArray);
    
    $objPHPExcel->getActiveSheet()->getStyle('A5:L5')->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'F2F2F2')
            )
        )
    );
    $objPHPExcel->getActiveSheet()->getStyle('M5')->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '00FF00')
            )
        )
    );
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("D4:G4");
    $objPHPExcel->getActiveSheet()->getStyle('D4:G4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
//    $objPHPExcel->getActiveSheet()->getStyle('D4:G4')->applyFromArray(
//        array(
//            'fill' => array(
//                'type' => PHPExcel_Style_Fill::FILL_SOLID,
//                'color' => array('rgb' => 'FDFDD9')
//            )
//        )
//    );
        
    
    
    $styleArray3 = array(
	'borders' => array(
		'outline' => array(
			'style' => PHPExcel_Style_Border::BORDER_MEDIUM			
		),
	),
    );
    $objPHPExcel->getActiveSheet()->getStyle("D4:G4")->applyFromArray($styleArray3);
    
    /** Saving the file * */
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    ob_end_clean();
    $objWriter->save('php://output');
}



function getOrderReports($date,$status="",$location="")
{
    global $conn,$objPHPExcel,$rowCount;
    $statusTxt = ($status == "")?"":" AND o.`status` = 'void'";
    if($status == "location" ) {
        $statusTxt = " AND o.`location_to` = '$location'";
    }

    $stmt = $conn->prepare('SELECT DISTINCT o.`order_id` as order_id,o.`date` as dt,o.`location_from` as loc_from,o.`location_to` as loc_to,o.`type` as typ,o.`status` as status,p.`product_name` as p_name,p.`qty_pallet` as pallet,p.`qty_case` as cas,p.`qty_bottle` as bottle,p.`qty_partial` as partial FROM `order` o, `bill` p WHERE DATE(o.`date`) between :dt1 and :dt2 AND o.`order_id` = p.`order_id`'.$statusTxt);
    $stmt->execute(array('dt1'=>$date['dateFrom'],'dt2'=>$date['dateTo']));    
    
    if($status == "")
    {
        $chargeable_location = getChargeableLocation();
        $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Order #');
        $objPHPExcel->getActiveSheet()->SetCellValue('B1', 'Date');
        $objPHPExcel->getActiveSheet()->SetCellValue('C1', 'Location To');
        $objPHPExcel->getActiveSheet()->SetCellValue('D1', 'Chargeable/Non Chargeable');
        $objPHPExcel->getActiveSheet()->SetCellValue('E1', 'Location From');
        $objPHPExcel->getActiveSheet()->SetCellValue('F1', 'Chargeable/Non Chargeable');
        $objPHPExcel->getActiveSheet()->SetCellValue('G1', 'Product Name');
        $objPHPExcel->getActiveSheet()->SetCellValue('H1', 'Pallet');
        $objPHPExcel->getActiveSheet()->SetCellValue('I1', 'Cases');
        $objPHPExcel->getActiveSheet()->SetCellValue('J1', 'Individuals');
        $objPHPExcel->getActiveSheet()->SetCellValue('K1', 'Partials');
        $objPHPExcel->getActiveSheet()->SetCellValue('L1', 'Type');
        $objPHPExcel->getActiveSheet()->SetCellValue('M1', 'Status');
    } else {
        $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Order #');
        $objPHPExcel->getActiveSheet()->SetCellValue('B1', 'Date');
        $objPHPExcel->getActiveSheet()->SetCellValue('C1', 'Location To');
        $objPHPExcel->getActiveSheet()->SetCellValue('D1', 'Location From');
        $objPHPExcel->getActiveSheet()->SetCellValue('E1', 'Product Name');
        $objPHPExcel->getActiveSheet()->SetCellValue('F1', 'Pallet');
        $objPHPExcel->getActiveSheet()->SetCellValue('G1', 'Cases');
        $objPHPExcel->getActiveSheet()->SetCellValue('H1', 'Individuals');
        $objPHPExcel->getActiveSheet()->SetCellValue('I1', 'Partials');
        $objPHPExcel->getActiveSheet()->SetCellValue('J1', 'Type');
        $objPHPExcel->getActiveSheet()->SetCellValue('K1', 'Status');
    }
    
        
    $styleArray = array(
      'borders' => array(
          'allborders' => array(
              'style' => PHPExcel_Style_Border::BORDER_MEDIUM
          )
      )
    );
    $col = ($status == "") ? "M" : "K";
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        if($status == "") {
            $charge1 = in_array($row['loc_to'],$chargeable_location['CHARGEABLE']) ? "Chargeable" : "Non Chargeable";
            $charge2 = in_array($row['loc_from'],$chargeable_location['CHARGEABLE']) ? "Chargeable" : "Non Chargeable";
            if($row['loc_to'] == "Boneyard") {
                $charge1 = "";
            }
            if($row['loc_from'] == "Boneyard") {
                $charge2 = "";
            }
            $objPHPExcel->getActiveSheet()->SetCellValue('A' . $rowCount, $row['order_id']);
            $objPHPExcel->getActiveSheet()->SetCellValue('B' . $rowCount, $row['dt']);
            $objPHPExcel->getActiveSheet()->SetCellValue('C' . $rowCount, $row['loc_to']);
            $objPHPExcel->getActiveSheet()->SetCellValue('D' . $rowCount, $charge1);
            $objPHPExcel->getActiveSheet()->SetCellValue('E' . $rowCount, $row['loc_from']);
            $objPHPExcel->getActiveSheet()->SetCellValue('F' . $rowCount, $charge2);
            $objPHPExcel->getActiveSheet()->SetCellValue('G' . $rowCount, $row['p_name']);
            $objPHPExcel->getActiveSheet()->SetCellValue('H' . $rowCount, $row['pallet']);
            $objPHPExcel->getActiveSheet()->SetCellValue('I' . $rowCount, $row['cas']);
            $objPHPExcel->getActiveSheet()->SetCellValue('J' . $rowCount, $row['bottle']);
            $objPHPExcel->getActiveSheet()->SetCellValue('K' . $rowCount, $row['partial']);
            $objPHPExcel->getActiveSheet()->SetCellValue('L' . $rowCount, $row['typ']);
            $objPHPExcel->getActiveSheet()->SetCellValue('M' . $rowCount, $row['status']);
        } else {
            $objPHPExcel->getActiveSheet()->SetCellValue('A' . $rowCount, $row['order_id']);
            $objPHPExcel->getActiveSheet()->SetCellValue('B' . $rowCount, $row['dt']);
            $objPHPExcel->getActiveSheet()->SetCellValue('C' . $rowCount, $row['loc_to']);
            $objPHPExcel->getActiveSheet()->SetCellValue('D' . $rowCount, $row['loc_from']);
            $objPHPExcel->getActiveSheet()->SetCellValue('E' . $rowCount, $row['p_name']);
            $objPHPExcel->getActiveSheet()->SetCellValue('F' . $rowCount, $row['pallet']);
            $objPHPExcel->getActiveSheet()->SetCellValue('G' . $rowCount, $row['cas']);
            $objPHPExcel->getActiveSheet()->SetCellValue('H' . $rowCount, $row['bottle']);
            $objPHPExcel->getActiveSheet()->SetCellValue('I' . $rowCount, $row['partial']);
            $objPHPExcel->getActiveSheet()->SetCellValue('J' . $rowCount, $row['typ']);
            $objPHPExcel->getActiveSheet()->SetCellValue('K' . $rowCount, $row['status']);
        }

        if( strtolower($row['status']) == 'void' && $statusTxt == "" ) {

            $objPHPExcel->getActiveSheet()->getStyle("A$rowCount:$col"."$rowCount")->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => 'FF3F3F')
                        )
                    )
                );
        } 
        
        $objPHPExcel->getActiveSheet()->getStyle("A$rowCount:$col"."$rowCount")->applyFromArray($styleArray);
        
        $objPHPExcel->getActiveSheet()->getStyle('A'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $rowCount++;
    }
    
    $objPHPExcel->getActiveSheet()->getStyle("A1:$col"."1")->getFont()->setBold(true);
    if($status == "")
    {
        $d = 26.00;
        $f = 26.00;
        $g = 21.00;
    } else {
        $d = 20.00;
        $f = 10.00;
        $g = 10.00;
    }
    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(15.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(22.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(18.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth($d);
    $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(30.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth($f);
    $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth($g);
    $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(12.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(10.00);

    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col.'1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
    $objPHPExcel->getActiveSheet()->getStyle("A1:".$col."1")->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => 'D8D8D8')
            )
        )
    );
    
    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col.'1')->applyFromArray($styleArray);
    /** Saving the file * */
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    ob_end_clean();
    $objWriter->save('php://output');
}

function getProductReport($date) {
    
        global $conn,$objPHPExcel,$rowCount;
    $stmt = $conn->prepare('SELECT DISTINCT o.`order_id` as order_id,o.`date` as dt,o.`location_from` as loc_from,o.`location_to` as loc_to,o.`type` as typ,o.`status` as status,p.`product_name` as p_name,p.`qty_pallet` as pallet,p.`qty_case` as cas,p.`qty_bottle` as bottle,p.`qty_partial` as partial FROM `order` o, `bill` p WHERE DATE(o.`date`) between :dt1 and :dt2 AND o.`order_id` = p.`order_id` ORDER BY p.`product_name`');
    $stmt->execute(array('dt1'=>$date['dateFrom'],'dt2'=>$date['dateTo']));    
    
    
    $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Order #');
    $objPHPExcel->getActiveSheet()->SetCellValue('B1', 'Date');
    $objPHPExcel->getActiveSheet()->SetCellValue('C1', 'Location To');
    $objPHPExcel->getActiveSheet()->SetCellValue('D1', 'Location From');
    $objPHPExcel->getActiveSheet()->SetCellValue('E1', 'Product Name');
    $objPHPExcel->getActiveSheet()->SetCellValue('F1', 'Pallet');
    $objPHPExcel->getActiveSheet()->SetCellValue('G1', 'Cases');
    $objPHPExcel->getActiveSheet()->SetCellValue('H1', 'Individuals');
    $objPHPExcel->getActiveSheet()->SetCellValue('I1', 'Partials');
    $objPHPExcel->getActiveSheet()->SetCellValue('J1', 'Type');
    $objPHPExcel->getActiveSheet()->SetCellValue('K1', 'Status');
    
        
    $styleArray = array(
      'borders' => array(
          'allborders' => array(
              'style' => PHPExcel_Style_Border::BORDER_MEDIUM
          )
      )
    );
    
    $nameCheck = "";
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
                        
        if($nameCheck != "" && $nameCheck != $row['p_name'])
        {
            $rowCount++;            
        }
        
        $objPHPExcel->getActiveSheet()->SetCellValue('A' . $rowCount, $row['order_id']);
        $objPHPExcel->getActiveSheet()->SetCellValue('B' . $rowCount, $row['dt']);
        $objPHPExcel->getActiveSheet()->SetCellValue('C' . $rowCount, $row['loc_to']);
        $objPHPExcel->getActiveSheet()->SetCellValue('D' . $rowCount, $row['loc_from']);
        $objPHPExcel->getActiveSheet()->SetCellValue('E' . $rowCount, $row['p_name']);
        $objPHPExcel->getActiveSheet()->SetCellValue('F' . $rowCount, $row['pallet']);
        $objPHPExcel->getActiveSheet()->SetCellValue('G' . $rowCount, $row['cas']);
        $objPHPExcel->getActiveSheet()->SetCellValue('H' . $rowCount, $row['bottle']);
        $objPHPExcel->getActiveSheet()->SetCellValue('I' . $rowCount, $row['partial']);
        $objPHPExcel->getActiveSheet()->SetCellValue('J' . $rowCount, $row['typ']);
        $objPHPExcel->getActiveSheet()->SetCellValue('K' . $rowCount, $row['status']);

        $nameCheck = $row['p_name'];
        if(strtolower($row['status']) == 'void') {
            
            $objPHPExcel->getActiveSheet()->getStyle("A$rowCount:K$rowCount")->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => 'FF3F3F')
                        )
                    )
                );
        } 
        $objPHPExcel->getActiveSheet()->getStyle("A$rowCount:K$rowCount")->applyFromArray($styleArray);
        
        $objPHPExcel->getActiveSheet()->getStyle('A'. $rowCount)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $rowCount++;
    }
    
    $objPHPExcel->getActiveSheet()->getStyle("A1:K1")->getFont()->setBold(true);
    
    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(22.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(30.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(11.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(10.00);
    $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(10.00);

    $objPHPExcel->getActiveSheet()->getStyle('A1:K1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
    $objPHPExcel->getActiveSheet()->getStyle("A1:K1")->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => 'D8D8D8')
            )
        )
    );
    
    $objPHPExcel->getActiveSheet()->getStyle('A1:K1')->applyFromArray($styleArray);
    /** Saving the file * */
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    ob_end_clean();
    $objWriter->save('php://output');
}

?>