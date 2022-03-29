<?php

session_start();
$con = mysqli_connect('localhost','root', '', 'generalsheet');
require '../vendor/autoload.php';

if(isset($_POST['save_excel_data'])){
    $fileName = $_FILES['import_file']['name'];
    $file_ext = pathinfo($fileName, PATHINFO_EXTENSION);

    $allowed_ext = ['xls', 'csv', 'xlsx'];

        if(in_array($file_ext, $allowed_ext)){
            $inputFileNamePath = ['import_file']['tmp_name'];
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileNamePath);
            $dta = $spreadsheet->getActiveSheet()->toArray();

            $count = "0";

            foreach($data as $row){
                if($count > 0){

                $Task = $row['0'];
                $Description = $row['1'];
                $Category = $row['2'];
                $priority = $row['3'];
                $EstimatedHours = $row['4'];
                $MVP = $row['5'];

                $generalsheetDBQuery = "INSERT INTO generalsheetDB (Task,Description,Category,priority,EstimatedHours,MVP) VALUES ('$Task','$Description','$Category','$priority','$EstimatedHours','$MVP')";
                $result = mysqli_query($con, $generalsheetDBQuery);
                $msg = true;
            }
            else{
                    $count = "1"
            }
        }

            if(isset($msg)){
                $_SESSION['message']= 'succuddfully imported';
                header('Location: index.php');
                exit(0);
            } else{
                $_SESSION['message']= 'Not imported';
                header('Location: index.php');
                exit(0);
            }
        }else{
            $_SESSION['message'] = "Invalid file";
            header('Location: index.php');
            exit(0);
        }

    }

?>