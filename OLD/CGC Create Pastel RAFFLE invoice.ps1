<#      Extract from Customer spreadsheet the range for new invoices to be generated.
        Modify the $startR (Startrow) if required.
        Modify the $endR (endrow) required.
#>
$inspreadsheet = 'C:\Users\Guy\Dropbox\CGC\2022\Training\10September 2022\Cape Gundog Club - Training Registration - 10th & 11th September 2022.xlsx'     #Spreadsheet from Booking System
$startR = 1                                     #Start row
$endR =  11                                   #End Row
$startCol = 1                                   #Start Col (don't change)
$endCol = 7                                    #End Col (don't change)
$csvfile = 'Filter1.csv'                                        
$csvfile2 = 'Filter2NHDR.csv'
$pastelfile = 'Pastel_Invoices.txt'                     #Invoice file to be imported into Pastel.
$pathout = 'C:\userdata\cgc local\Invoices\'
$custsheet = 'Invoice'                           #Customer worksheet IN $inspreadsheet (do not change)
#
$Outfile = $pathout + $csvfile
$Outfile2 = $pathout + $csvfile2
$Outfile4 = $pathout + $pastelfile
#
Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol | Where-Object -FilterScript {$_.AccCode -like "*" -and $_.Generate -eq "Yes"} | Export-CSV -Path $Outfile -NoTypeInformation
#
Get-Content -Path $Outfile | Select-Object -skip 1 | Set-Content -path $Outfile2            #Remove Row header
Remove-Item -Path $outfile
<#  
    Get list of invoices from spreadsheet
    Output to text file to be imported as a Pastel Invoice batch.
#>
$outfile3 = 'c:\userdata\cgc local\invoices\WMinvTmp.txt'                 #
#Remove last file imported to Pastel
if (Test-Path $outfile3) { Remove-Item $outfile3 }
#Import latest csv from Client spreadsheet

$invoicedate = (Get-Date -UFormat "%d/%m/%Y")

$data = Import-Csv -path $Outfile2 -header accnum, name, dog, amt, event, eventdate

$prevaccnum = 0

foreach ($aObj in $data) {
    
    if ($aObj.accnum -eq "") {Break} 
    
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $invoicedate
    
    Switch ($aObj.event) {
        ACCOM {$incomeAcc = '1550000'}
        FT {$incomeAcc = '1000000'}
        LIMIT {$incomeAcc = '1000000'}
        MEALS {$incomeAcc = '1200010'}
        RAFFLE {$incomeAcc = '1300000'}
        LANYARD {$incomeAcc = '1300000'}
        REFRESHMENTS {$incomeAcc = '1200011'}
        TRAINING {$incomeAcc = '1400000'}
    }
    # If booking id has changed then add a header record,
    # this with happen for the first header as well
    if ($aObj.accnum -ne $prevaccnum) {
        $prevaccnum = $aObj.accnum

        $headerProperties = [ordered] @{
            Col1  = 'Header'
            Col2  = ''
            Col3  = ''
            Col4  = 'Y'
            Col5  = $aObj.accnum
            Col6  = $pastelper
            Col7  = $invoicedate
            Col8  = ""
            Col9  = "N"
            Col10 = '0'
            Col11 = ''
            Col12 = ''
            Col13 = ''
            Col14 = ''
            Col15 = ''
            Col16 = ''
            Col17 = ''
            Col18 = ''
            Col19 = ''
            Col20 = '0'
            Col21 = $invoicedate
            Col22 = ''
            Col23 = ''
            Col24 = ''
            Col25 = '1'
            Col26 = ''
            Col27 = ''
            Col28 = ''
            Col29 = ''
        }
        $Line1Properties = [ordered] @{    
            Col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            Col10 = "'"
            Col11 = 'Handler'  +  ' : ' + $aObj.name
            Col12 = 7
            Col13 = ''
            Col14 = ''
            Col15 = ''
            Col16 = ''
            Col17 = ''
            Col18 = ''
            Col19 = ''
            Col20 = ''
            Col21 = ''
            Col22 = ''
            Col23 = ''
            Col24 = ''
            Col25 = ''
            Col26 = ''
            Col27 = ''
            Col28 = ''
            Col29 = '' 
        }
        #
        $Line2Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = ''
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
        $Line3Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = $aObj.event + ' - ' + $aObj.eventdate
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
       <# $Line4Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = ''
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
        $Line5Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = ''
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
        $Line6Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = ''
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
        $Line7Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = ' '
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
        #>
        # Append the header and invoice lines to the CSV file
        $objHeader = New-Object -TypeName psobject -Property $headerProperties 
        $objHeader | Select-Object * | Export-Csv -path $outfile3 -Append -NoTypeInformation

        $objGroup = New-Object -TypeName psobject -Property $Line1Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile3 -Append -NoTypeInformation

        $objGroup = New-Object -TypeName psobject -Property $Line2Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile3 -Append -NoTypeInformation
        
        $objGroup = New-Object -TypeName psobject -Property $Line3Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile3 -Append -NoTypeInformation
        <#
        $objGroup = New-Object -TypeName psobject -Property $Line4Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile3 -Append -NoTypeInformation
        
        $objGroup = New-Object -TypeName psobject -Property $Line5Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile3 -Append -NoTypeInformation
        
        $objGroup = New-Object -TypeName psobject -Property $Line6Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile3 -Append -NoTypeInformation
        
        $objGroup = New-Object -TypeName psobject -Property $Line7Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile3 -Append -NoTypeInformation
    #>    
    }
    #Add the current row to the objects
    $detailProperties = [ordered] @{
        Col1  = 'Detail'
        Col2  = '0'
        Col3  = '1'
        Col4  = $aObj.amt
        Col5  = $aObj.amt
        Col6  = ''
        Col7  = ''
        Col8  = '0'
        Col9  = '0'
        Col10 = $incomeAcc                    #which income account ?
        Col11 = $aObj.dog
        Col12 = '6'
        Col13 = ''
        Col14 = ''
        Col15 = ''
        Col16 = ''
        Col17 = ''
        Col18 = ''
        Col19 = ''
        Col20 = ''
        Col21 = ''
        Col22 = ''
        Col23 = ''
        Col24 = ''
        Col25 = ''
        Col26 = ''
        Col27 = ''
        Col28 = ''
        Col29 = ''
    } 

    $objDetails = New-Object -TypeName psobject -Property $detailProperties 
    $objDetails | Select-Object * | Export-Csv -path $outfile3 -Append -NoTypeInformation
}  
#Remove header information so file can be imported into Pastel Accounting.

Get-Content -Path $outfile3 | Select-Object -skip 1 | Set-Content -path $outfile4
Remove-Item -Path $outfile3
Remove-Item -Path $Outfile2