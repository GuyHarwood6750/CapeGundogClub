<#      Extract credit card transactions from spreadsheet to be processed as Pastel payment batch.
        Modify the $endR (endrow) as spreadsheet is added too.

#>
$inspreadsheet = 'C:\Users\Guy\Dropbox\CGC\2024\Financials\Bank Statements\ABSA Bank Transactions.xlsx'
$outfile2 = 'C:\Users\Guy\Dropbox\CGC\2024\Financials\Bank Statements\CHQ_Transactions_1.csv'
$custsheet = '7348 Transactions'                                #Transactions worksheet
$startR = 2                                         #Start row
$endR = 5                                              #End Row - change if necessary depending on number of transactions
$csvfile = 'SHEET1.csv'
$pathout = 'C:\Users\Guy\Dropbox\CGC\2024\Financials\Bank Statements\'
$startCol = 1                                                                   #Start Col (don't change)
$endCol = 12                                                                     #End Col (don't change)
$filter = "R"                                                                   #P=Payments, R=Receipts\deposits
$outfile1 = 'C:\Users\Guy\Dropbox\CGC\2024\Financials\Bank Statements\CHQTEMP.txt'              #Temp file
$outfileF = 'C:\Users\Guy\Dropbox\CGC\2024\Financials\Bank Statements\7348_Transactions_pastel_' + $filter + '.txt'  #File to be imported into Pastel             
$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader -DataOnly| Where-Object -Filterscript { $_.P1 -eq $filter -and $_.P11 -ne 'done'} | Export-Csv -Path $Outfile -NoTypeInformation

ExcelFormatDate -file $Outfile -sheet 'SHEET1' -column 'D:D'

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile

#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfilef }                   

#Import latest csv from Client spreadsheet
$data = Import-Csv -path $outfile2 -header type, GL, Expacc, date, ref, date2, desc, amt, bal, amt1, vat     

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods2 -transactiondate $aObj.date
    
    Switch ($aObj.Expacc) {
        AMMO { $expacc = '2000004'; $aObj.desc }         
        FEE { $expacc = '3200000'; $aObj.desc }         
        INTR { $expacc = '2750000'; $aObj.desc }
        SGREEN { $expacc = $aObj.Expacc; $aObj.descr }               #Customer                 
        STAV { $expacc = $aObj.Expacc; $aObj.descr }               #Customer                 
        LDIME { $expacc = $aObj.Expacc; $aObj.descr }               #Customer                 
        #CON001 { $expacc = $aObj.Expacc; $aObj.descr }              #Supplier         
        
        #Default { $expacc = '9999000'; $aObj.desc }                  
        Default { $expacc = $aObj.Expacc; $aObj.desc }                  
    }

    Switch ($aObj.vat) {
        Y { $VATind = '15' }
        N { $VATind = '0' }
        Default {$VATind = '15'}
    }
    #Format Pastel batch   
    $props1 = [ordered] @{
        Period  = $pastelper
        Date    = $aObj.date
        GL      = $aObj.GL                      #GDC - general ledger, debtor, creditor
        contra  = $expacc                       #Expense account to be debited (DR)
        ref     = $aObj.ref
        comment = $aObj.desc
        amount  = $aObj.amt1
        fil1    = $VATind
        fil2    = '0'
        fil3    = ' '
        fil4    = '     '
        fil5    = '8420000'                     #Cheque account contra account number
        fil6    = '1'
        fil7    = '1'
        fil8    = '0'
        fil9    = '0'
        fil10   = '0'
        amt2    = $aObj.amt1
    }
      
        $objlist = New-Object -TypeName psobject -Property $props1
        $objlist | Select-Object * | Export-Csv -path $outfile1 -NoTypeInformation -Append
    }  
    #Remove header information so file can be imported into Pastel Accounting.
    Get-Content -Path $outfile1 | Select-Object -skip 1 | Set-Content -path $outfilef
    Remove-Item -Path $outfile1