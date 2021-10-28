#PaperCut Print job Report
#Built By Alexis Daigle
#CSVFIX is needed to run this https://code.google.com/p/csvfix/downloads/list
#This Script will colect all the csv files created by PaperCut print loger and create a per month\per user usage report.
#This scripts will use a folder call Plogs to work the data, create this folder or change it in the script.


#get Month and year of the report wanted
$Year=Read-Host "Please enter a Year ex:2014"
$Month=Read-Host "Please enter a Month ex:03"
$filename
#reset values for counters
$all = 0
$PC = 0
$PageTotal = 0
"You picked $Year $Month"
#copy monthly report to Local Machine
cp "\\Server1\C$\Program Files (x86)\PaperCut Print Logger\logs\csv\monthly\papercut-print-log-$Year-$Month.csv" "C:\Plogs\S1-$Year-$Month-raw.csv"
cp "\\Server2\C$\Program Files (x86)\PaperCut Print Logger\logs\csv\monthly\papercut-print-log-$Year-$Month.csv" "C:\Plogs\S2-$Year-$Month-raw.csv"
cp "\\Server3\C$\Program Files\PaperCut Print Logger\logs\csv\monthly\papercut-print-log-$Year-$Month.csv" "C:\Plogs\S3-$Year-$Month-raw.csv"
#Removes the first rows of Vendor tags which csvs dont like. To be powershell only if posible on a later date.
& 'C:\Program Files (x86)\CSVfix\csvfix.exe' remove -if '$line <2' -o "C:\Plogs\S1-$Year-$Month.clr.csv" "C:\Plogs\S1-$Year-$Month-raw.csv"
& 'C:\Program Files (x86)\CSVfix\csvfix.exe' remove -if '$line <2' -o "C:\Plogs\S2-$Year-$Month.clr.csv" "C:\Plogs\S3-$Year-$Month-raw.csv"
& 'C:\Program Files (x86)\CSVfix\csvfix.exe' remove -if '$line <2' -o "C:\Plogs\S3-$Year-$Month.clr.csv" "C:\Plogs\S3-$Year-$Month-raw.csv"
#delete Raw data
rm "C:\Plogs\S1-$Year-$Month-raw.csv"
rm "C:\Plogs\S2-$Year-$Month-raw.csv"
rm "C:\Plogs\S3-$Year-$Month-raw.csv"
#Import CSV's, Merge, Sort and Export
$PRTCSV=Import-Csv "C:\Plogs\S1-$Year-$Month.clr.csv" 
$P64CSV=Import-Csv "C:\Plogs\S2-$Year-$Month.clr.csv"
$P32CSV=Import-Csv "C:\Plogs\S3-$Year-$Month.clr.csv"
$all=$PRTCSV+$P64CSV+$P32CSV
#delete Clean data
rm "C:\Plogs\S1-$Year-$Month.clr.csv"
rm "C:\Plogs\S2-$Year-$Month.clr.csv"
rm "C:\Plogs\S3-$Year-$Month.clr.csv"
#$all | Sort-Object Time | Export-Csv "c:\Plogs\Final$Year-$Month.csv" -NoTypeInformation
#"Unique Print Jobs Per Printer For $Year / $Month" > c:\Plogs\Print-$Year-$Month.txt
$Print=($all | Group-Object Printer -NoElement | Sort-Object Count -Descending ) 
#$Print >> c:\Plogs\Print-$Year-$Month.txt
#$Print | Out-GridView
#"Unique Print Jobs Per User For $Year / $Month" > c:\Plogs\User-$Year-$Month.txt
#Create a new text File to contain the Report add Percentage.(function to calculate total page count then for each printer)
#Write-Host "Unique,Printer,PageCount,Percentage"
"Unique,Printer,PageCount,Percentage" > c:\Plogs\Final$Year$Month.txt
#Function to Sum up the total page count for percentage. $PCT=PageCountTotal
"Calulating Total Page Count"
foreach ($ROW in $all ){
		
			$ROWP = $ROW.Pages
			$ROWC = $ROW.Copies
			$PageTotal = [int]$PageTotal + (([int]$ROWP)*([int]$ROWC))
		}
#Function to sum up all the print jobs for each printer.
"Calulating Printer Usage"
foreach ($Printer in $Print){
	$Prt = $Printer.Name
	$Unique = $Printer.Count
	foreach ($PT in $all ){
		if ($PT.Printer -eq $Prt){
			$PTP = $PT.Pages
			$PTC = $PT.Copies
			$PC = [int]$PC + (([int]$PTP)*([int]$PTC))
		}
		Else {}
	}
	#calc the Percentage, format to proper percentage value.
	$percent = ([int]$PC / [int]$PageTotal)
	$percentV = "{0:P}" -f $Percent
	$UniqueT += $Unique
	#append new data to new CSV file, reset page count and $percentage to 0.
	#Write-Host "$Unique,$Prt,$PC,$percentV"
	"$Unique,$Prt,$PC,$percentV" >> c:\Plogs\Final$Year$Month.txt
	$PC = 0
	$percent = 0
	}
"$UniqueT,Totals,$PageTotal,####" >> c:\Plogs\Final$Year$Month.txt
$UniqueT = 0
$Print = 0
"Unique,User,PageCount,Percentage" >> c:\Plogs\Final$Year$Month.txt
$users=($all | Group-Object User -NoElement | Sort-Object Count -Descending )
#Function to sum up all the print jobs for each users.
"Calulating User Printage"
foreach ($user in $users){
	$Usern = $user.Name
	$Unique = $user.Count
	
	foreach ($Use in $all ){
		if ($Use.User -eq $Usern){
			$PTP = $Use.Pages
			$PTC = $Use.Copies
			$PC = [int]$PC + (([int]$PTP)*([int]$PTC))
			}
		Else {}
	}
	#calc the Percentage, format to proper percentage value.
	$percent = ([int]$PC / [int]$PageTotal)
	$percentV = "{0:P}" -f $Percent
	$UniqueT += $Unique
	#append new data to new CSV file, reset page count and $percentage to 0.
	#Write-Host "$Unique,$Usern,$PC,$percentV"
	"$Unique,$Usern,$PC,$percentV" >> c:\Plogs\Final$Year$Month.txt
	$PC = 0
	$percent = 0
	}
"$UniqueT,Totals,$PageTotal,####" >> c:\Plogs\Final$Year$Month.txt
$PageTotal = 0
$UniqueT = 0
#check if old version is still located there, delete the old one.
if(Test-Path "c:\Plogs\Final$Year$Month.csv"){
	rm "c:\Plogs\Final$Year$Month.csv" }
Import-Csv -Path c:\Plogs\Final$Year$Month.txt | Export-Csv -Path c:\Plogs\Final$Year$Month.csv -NoTypeInformation
rm "c:\Plogs\Final$Year$Month.txt"
"Finished, Report can be found in C:\Plogs"
