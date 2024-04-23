Param(
    
    # Путь к файлу реестров. Если нужно просто сделать заявку, просто не заполнять.
    # Файл должны быть сохранён в кодировке utf-8, т.к. Select-String не воспринимает другие кодировки на момент разработки сценария.
    [ValidateScript( { -not $FilePath -or ( (Test-Path $_) -and ( $_.EndsWith('.txt') ) ) } )]
    [string]$FilePath,
    
    [ValidateScript( { ( (Test-Path $_) -and ( $_.EndsWith('.txt') ) ) } )]
    [string]$VinPath,
    
    [ValidateScript( { $_ -match '\d{9}' } )]
    [string]$OrganizationINN = '',
    
    [string]$OrganizationName = '',
    
    [string]$OrganizationAddr = '',
    
    [string]$OrganizationEmail = '',

    [Parameter(Mandatory=$true)]
    [string]$ConnectionString,

    [Parameter(Mandatory=$true)]
    [string]$DBname = '',
    
    [ValidateScript( { ( ( Test-Path $_ ) -and ( $_.EndsWith('.doc') -or $_.EndsWith('.docx') ) ) } )]
    [string]$WordApplicationRequestDocPath = "${ENV:USERPROFILE}\Documents\Обработка данных реестров ДСП и ЦНИИТУ\Исправленная заявка ЭПТС.docx",
 
    [switch]$DoNotShowWordApplication,
 
    # Если не нужно добавлять только по VIN с отражёнными в 1С счетами-фактурами
    [switch]$WithoutInvoices
 
)
 
$filterVIN = (Get-Content $VinPath).Split() | Where-Object { $_ }
$lettersTable = @()
 
if ( $FilePath ) {
 
    [regex]$rxLetter = '^\s*№.+\s+от\s+\d\d\.\d\d\.\d\d\d\d\s*[\.г]?\s*$'
    [regex]$rxNpp = '^\s*\d+\s*$'
 
    $letterCatches = Select-String -Pattern $rxLetter -Path $FilePath
    $nppCatches = Select-String -Pattern $rxNpp -Path $FilePath
 
}
 
forEach($VIN in $filterVIN)
{
    if ( $FilePath ) {
 
        $curCatch = Select-String -Pattern $VIN -Path $FilePath
 
        if ( $curCatch )
        {
            $curLetterCatch = $letterCatches | Where-Object { $_.LineNumber -lt $curCatch.LineNumber } | Sort-Object -Property LineNumber | Select-Object -Last 1
            $curNppCatch = $nppCatches | Where-Object { $_.LineNumber -lt $curCatch.LineNumber } | Sort-Object -Property LineNumber | Select-Object -Last 1
 
            $lettersTable += 1 | Select-Object  @{ Name = 'Number';       Expression = { $curNppCatch.Line.Trim() } },
                                                @{ Name = 'Letter';       Expression = { $curLetterCatch.Line.Trim() } },
                                                @{ Name = 'VIN';          Expression = { $VIN } }
        }
 
    } else {
 
        $lettersTable += 1 | Select-Object  @{ Name = 'Number';       Expression = { $null } },
                                            @{ Name = 'Letter';       Expression = { ' ' } },
                                            @{ Name = 'VIN';          Expression = { $VIN } }
 
    }
}
 
if ( $lettersTable )
{
    
    $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $connectionString
    $connection.Open()
 
    if ($connection.State -eq 'Open') {
        
        if ( -not $WithoutInvoices ) {
        $sql = @"
select s._Description Series
, case e._EnumOrder
    when 0 then 'Электронная ТН-2'
    when 1 then 'Электронная ТТН-1'
    when 2 then 'ТТН-1'
    when 3 then 'ТН-2'
    when 4 then 'Акт выполненных работ'
    when 5 then 'Счет-фактура'
    when 6 then 'CMR-накладная'
    when 7 then 'Invoice (счет)'
    when 8 then 'Авизо'
    when 9 then 'Акт'
    when 10 then 'Бухгалтерская справка'
    when 11 then 'Договор'
    when 12 then 'Другое'
    when 13 then 'Коносамент'
    when 14 then 'Контракт'
else 'н/у'
end DocType
, d._Fld45746 IncomeDocSeries
, d._Fld23064 IncomeDocNumber
, FORMAT(d._Fld23065, 'dd.MM.yyyy') IncomeDocDate
, g._Description GoodCommercialName
, s._Description VIN
, agr._Description AgreementName
, FORMAT(agr._Fld3613, 'dd.MM.yyyy') AgreementDate
from $DBname.dbo._Document694_VT23098 dg
inner join $DBname.dbo._Document694 d
on d._IDRRef = dg._Document694_IDRRef
inner join $DBname.dbo._Reference223 o
on d._Fld23041RRef = o._IDRRef
left join $DBname.dbo._Reference126 agr
on d._Fld23088RRef = agr._IDRRef
inner join $DBname.dbo._Reference351 s
on dg._Fld23131RRef = s._IDRRef
left join $DBname.dbo._Enum45639 e
on e._IDRRef = d._Fld45747RRef
left join $DBname.dbo._Reference209 g
on dg._Fld23100RRef = g._IDRRef
where
d._Posted > 0
AND o._Fld5664 = @INN
AND s._Description IN (
"@
        } else {
            $sql = @"
select max(s._Description) Series
    , 'Контракт' DocType
    , max(ISNULL(g._Description, models._Description)) GoodCommercialName
    , s._Description VIN
    , max(agr._Description) AgreementName
    , max(FORMAT(DATEADD(year, -2000, agr._Fld3613), 'dd.MM.yyyy')) AgreementDate
from
$DBname.dbo._Reference351 s
    left join $DBname.dbo._Document694_VT23098 dg
    on dg._Fld23131RRef = s._IDRRef
    left join $DBname.dbo._Document694 d
    on d._IDRRef = dg._Document694_IDRRef
        and d._Posted > 0
    left join $DBname.dbo._Reference223 o
    on d._Fld23041RRef = o._IDRRef
        and o._Fld5664 = @INN
    left join $DBname.dbo._Reference126 agr
    on d._Fld23088RRef = agr._IDRRef
    left join $DBname.dbo._Enum45639 e
    on e._IDRRef = d._Fld45747RRef
        and e._EnumOrder = 14
    left join $DBname.dbo._Reference209 g
    on dg._Fld23100RRef = g._IDRRef
    left join $DBname.dbo._Reference209 models
    on s._Fld46127RRef = models._IDRRef
where
    s._Description IN (
"@
        }
        
        $lettersTable | ForEach-Object { $sql += "'" + $_.VIN.Replace(' ','').Replace(';','').Replace(',','').Replace('.','').Replace('(','').Replace(')','').Replace('@','') + "'," }
        $sql += ')'
        $sql = $sql.Replace(',)',')')
        if ($WithoutInvoices) {
            $sql += 'group by s._Description'
        }
 
        $command = New-Object -TypeName System.Data.SqlClient.SqlCommand $sql, $connection -ErrorAction Stop
        $command.Parameters.Add('INN', $OrganizationINN) | Out-Null
 
        $adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
        $table = New-Object -TypeName System.Data.DataTable
 
        $adapter.Fill($table)
        
        $connection.Close()
        $connection.Dispose()
 
    }
}
 
if ( $lettersTable -and $WordApplicationRequestDocPath )
{
    
    $dspLettersUnique = if ($FilePath) { ($lettersTable | Select-Object -Unique Letter).Letter } else { @(' ') }
 
    $registerNumber = 0
 
    foreach($currentLetter in $dspLettersUnique) {
 
        $registerNumber++
 
        $dspLetterVINs = ($lettersTable | Where-Object { $_.Letter -eq $currentLetter } | Select-Object -Unique VIN).VIN
 
        $WordApplication = New-Object -ComObject "Word.Application"
        $WordApplication.Documents.Open( $WordApplicationRequestDocPath ) | Out-Null
 
        $WordApplication.ActiveDocument.SaveAs2($WordApplication.ActiveDocument.FullName.Split('.')[-2] + "($registerNumber)") # + $WordApplication.ActiveDocument.FullName.Split('.')[-1])
    
        if (-not $DoNotShowWordApplication ) {
            $WordApplication.ShowStartupDialog = $true
            $WordApplication.Visible = $true
            $WordApplication.Activate()
        }
 
        $TableOne  = $WordApplication.ActiveDocument.Tables[1]
 
        $TableOne.Cell(1, 2).Range.Text = "$($OrganizationName)"
        $TableOne.Cell(2, 2).Range.Text = "$($OrganizationINN)"
        $TableOne.Cell(3, 2).Range.Text = "$($OrganizationAddr)"
        $TableOne.Cell(4, 2).Range.Text = "$($OrganizationEmail)"
 
        $TableTwo  = $WordApplication.ActiveDocument.Tables[3]
 
        $TableFour = $WordApplication.ActiveDocument.Tables[6]
 
        $TableFour.Cell(2, 4).Range.Text = "$($currentLetter)"
 
        $TableFive  = $WordApplication.ActiveDocument.Tables[7]
    
        if ( $table -or $WithoutInvoices ) {
 
            $currentRowNumberTbl5 = 2
            
            if ( $FilePath) {
                $table | Where-Object { $_.VIN -in $dspLetterVINs } | Select-Object -Unique AgreementName, AgreementDate | ForEach-Object {
 
                    if ( $currentRowNumberTbl5 -gt $TableFive.Rows.Count ) {
                        $TableFive.Rows.Add() | Out-Null
                    }
 
                    $TableFive.Cell($currentRowNumberTbl5, 1).Range.Text = $currentRowNumberTbl5 - 1
                    $TableFive.Cell($currentRowNumberTbl5, 2).Range.Text = "Контракт"
                    $TableFive.Cell($currentRowNumberTbl5, 3).Range.Text = "$($_.AgreementName)"
                    $TableFive.Cell($currentRowNumberTbl5, 4).Range.Text = "$($_.AgreementDate.ToString().Replace('.40','.20'))"
 
                    $currentRowNumberTbl5++
 
                }
            }
            
            $currentRowNumber = 2
 
            $newRows = $table | Where-Object { $_.VIN -in $dspLetterVINs }
 
            forEach($row in $newRows ) {
 
 
                if ( $currentRowNumber -gt $TableTwo.Rows.Count ) {
                    $TableTwo.Rows.Add() | Out-Null
                }
 
 
                if ( -not $WithoutInvoices -and $currentRowNumberTbl5 -gt $TableFive.Rows.Count ) {
                    $TableFive.Rows.Add() | Out-Null
                }
 
 
                $TableTwo.Cell($currentRowNumber, 1).Range.Text = $currentRowNumber - 1
                $TableTwo.Cell($currentRowNumber, 2).Range.Text = "$($row.GoodCommercialName)"
                $TableTwo.Cell($currentRowNumber, 4).Range.Text = "$($row.VIN)"
 
                if ( -not $WithoutInvoices ) {
                    $TableFive.Cell($currentRowNumberTbl5, 1).Range.Text = $currentRowNumberTbl5 - 1
                    $TableFive.Cell($currentRowNumberTbl5, 2).Range.Text = "$($row.DocType)"
                    $TableFive.Cell($currentRowNumberTbl5, 3).Range.Text = "$($row.IncomeDocSeries)$($row.IncomeDocNumber)"
                    $TableFive.Cell($currentRowNumberTbl5, 4).Range.Text = "$($row.IncomeDocDate.ToString().Replace('.40','.20'))"
                    
                    $currentRowNumberTbl5++
                }
 
                $currentRowNumber++
 
            }
        }
 
        $WordApplication.ActiveDocument.Save() | Out-Null
        
        if ( $DoNotShowWordApplication ) {
            $WordApplication.ActiveDocument.Close()
            $WordApplication.Quit()
        }
 
        $WordApplication = $null
 
    }
 
} elseif ( -not $lettersTable -and $WordApplicationRequestDocPath ) {
        
    Write-Warning "Не найдено ни одной подходящей строки. Проверьте параметры отбора."
 
} else {
    
    if ( $table ) {
    
        @{ Name = 'letters'; Expression = { $lettersTable } },
        @{ Name = 'incomeDocs'; Expression = { $table } }
    
    } else {
 
        $lettersTable
    
    }
 
}
