#Open 2 excel files and compares Selected columns and display the fields found in the 2nd column

#Author: Ernesto Farias 2019  - ernestfarias@gmail.com

cls
#SET PARAMETERS, no worries about file order it will compare the Smaller into the Larger
$FILE1="C:\Users\ernf\Documents\Report\FILE1.xlsx"
$WORKSHEET1="endpoints"
$COLNAME1="MACAddress"

$FILE2="C:\Users\ernf\Documents\Report\FILE2.xlsx"
$WORKSHEET2="endpoints"
$COLNAME2="MACAddress"

$SEARCHRANGE="" # "A:Z" look into the entire sheet, left empty ("") to look just in the same ColName column
$STARTPOSTION=1 #1 to avoid header comparison
 
#OPEN EXCEL AND WORKSHEETS
$Excel1 = New-Object -ComObject Excel.Application
$workbook1 = $Excel1.Workbooks.Open($FILE1)
$sheet1 = $workbook1.worksheets.Item($WORKSHEET1)

$Excel2 = New-Object -ComObject Excel.Application
$workbook2 = $Excel2.Workbooks.Open($FILE2)
$sheet2 = $workbook2.worksheets.Item($WORKSHEET2)

#return colum location index, assume are in 1st row between A and Z
$ColIndex1 = $sheet1.Range("a1:z1").Find($COLNAME1).Column
$ColIndex2 = $sheet2.Range("a1:z1").Find($COLNAME2).Column

$ColIndex1Name = $sheet1.Range("a1:z1").Find($COLNAME1).Address($false,$false)
$ColIndex2Name = $sheet2.Range("a1:z1").Find($COLNAME2).Address($false,$false)


echo "Column:$COLNAME1 Found in postion: $ColIndex1"
echo "Column:$COLNAME2 Found in postion: $ColIndex2"
$sheet1Length = ($sheet1.UsedRange.Rows).count
$sheet2Length = ($sheet2.UsedRange.Rows).count

echo "Length Sheet1=$sheet1Length"
echo "Length Sheet2=$sheet2Length`n"


#ITERATE ROWS
#clean vars
$i=$STARTPOSTION
$iter=""
$col1Val=""


#Choose smaller sheet length into the larger, so then comparing is more efficient

if ($sheet1Length -le $sheet2Length) {
$smallerSheet = $sheet1
$smallerSheetCol = $ColIndex1
$smallerSheetLength = $sheet1Length
$smallerSheetColIndexName = $ColIndex1Name

$largerSheet = $sheet2
$largerSheetCol = $ColIndex2
$largerSheetLength = $sheet2Length
$largerSheetColIndexName = $ColIndex2Name
echo "Compare values from File 1(Sheet:$($smallerSheet.Name) Column:$smallerSheetCol) to File 2($largerSheet.Name Column:$largerSheetCol)"
    } else {
    $smallerSheet = $sheet2
    $smallerSheetCol = $ColIndex2
    $smallerSheetLength = $sheet2Length
    $smallerSheetColIndexName = $ColIndex2Name

    $largerSheet = $sheet1
    $largerSheetCol = $ColIndex1
    $largerSheetLength = $sheet1Length
    $largerSheetColIndexName = $ColIndex1Name
    echo "Compare values from File 2 (Sheet:$($smallerSheet.Name) Column:$smallerSheetCol) to File 1 (Sheet:$($largerSheet.Name) Column:$largerSheetCol)"
}

sleep 2

$NotFounds=@()
$Founds=@()

while($i -le ($smallerSheet.UsedRange.Rows).count) 
    {
    $i++ 
        $col1Val = $smallerSheet.Cells.Item($i,$smallerSheetCol).text
        Write-Progress -Activity "Search in Progress" -Status "$i% Complete:" -PercentComplete $($smallerSheet.UsedRange.Rows).count;
        #for each new line in col1val, if new line is a list of lines, split it

          #replace newline tab space and cr by comma
          $pcol1val= $col1Val.replace("`n",",").replace("`r",",").replace("`t",",").replace(" ",",")
          foreach($val in $pcol1val.Split(",")){
  
          #HERE FIND EACH VAL IN OTHER EXCEL USED .FIND FUNTION
          $valst=$val -join " " #convert to string remove spaces
        #  echo "Value=$valst Len:$($valst.length)"

          if ($SEARCHRANGE -eq $null -or $SEARCHRANGE -eq "" )
          {$Range = $largerSheet.Range("$largerSheetColIndexName").EntireColumn}else{$Range = $largerSheet.Range("A:Z").EntireColumn}

          $Search = $Range.find($valst)
            if ($search -ne $null) {
                        $FirstAddress = $Search.Address
                        do {
                        #echo "FOUND in $($Search.Address($false,$false))"
                        $Founds += "$valst in $($Search.Address($false,$false))"
                        $Search = $Range.FindNext($Search)
                            } while ($Search -eq $null)#while ( $search -ne $null -and $search.Address -ne $FirstAddress)                
                } else {
         #       echo "NOTFound"
                $NotFounds += "$valst"}
}
}

#list not founds

echo "NOT FOUND:$($NotFounds.Count)"
foreach ($val in $NotFounds) {echo "$val"}

echo "`nFOUND ITEMS:$($Founds.Count)"
foreach ($val in $Founds) {echo "$val"}

$Excel1.quit()
$Excel2.quit()
