#Install-Module -Name ImportExcel
#Install-Module -Name ImportExcel -Scope CurrentUser -Verbose

$ImportExcel = Import-Excel -Path D:\Project\Programs\Powershell\Base.xlsx -WorksheetName "Mes Numéros"
$ResultLotoNumber = Import-Excel -Path D:\Project\Programs\Powershell\Base.xlsx -WorksheetName "Tirage-Loto"
$ResultEuroMillionsNumber = Import-Excel -Path D:\Project\Programs\Powershell\Base.xlsx -WorksheetName "Tirage-EuroMillions"


Foreach($values in $ImportExcel) 
{

    $MyArray = @{
        Source=$values.Source;First=$values.'1er n°';Second=$values.'2eme n°';Third=$values.'3eme n°';Forth=$values.'4eme n°';Fifth=$values.'5eme n°';Comp1=$values.'n° chance 1';Comp2=$values.'n° chance 2'
    }

    $Source = $values.Source
    $FirstMyNumber = $MyArray.First
    $SecondtMyNumber = $MyArray.Second
    $ThirdMyNumber = $MyArray.Third
    $ForthMyNumber = $MyArray.Forth
    $FifthMyNumber = $MyArray.Fifth
    $CompMyNumber1 = $MyArray.Comp1
    $CompMyNumber2 = $MyArray.Comp2

        If ($Source  -match "Loto")
        {
            $FirstMyNumberLoto = $FirstMyNumber
            $SecondtMyNumberLoto = $SecondtMyNumber
            $ThirdMyNumberLoto = $ThirdMyNumber
            $ForthMyNumberLoto = $ForthMyNumber
            $FifthMyNumberLoto = $FifthMyNumber
            $CompMyNumberLoto = $CompMyNumber1
            $MyArrayLoto = @($FirstMyNumberLoto,$SecondtMyNumberLoto,$ThirdMyNumberLoto,$ForthMyNumberLoto,$FifthMyNumberLoto)
            $MyArrayLotoComps = @($CompMyNumberLoto)
            $MyArrayLotoAll = @($FirstMyNumberLoto,$SecondtMyNumberLoto,$ThirdMyNumberLoto,$ForthMyNumberLoto,$FifthMyNumberLoto,$CompMyNumberLoto)

                Foreach($elementLoto in $ResultLotoNumber) 
                {

                $DayLoto =  $elementLoto.Jour
                $MonthLoto =  $elementLoto.Mois
                $YearLoto =  $elementLoto.Année
                $FirstNumberLoto = $elementLoto.'1er n°'
                $SecondtNumberLoto = $elementLoto.'2eme n°'
                $ThirdNumberLoto = $elementLoto.'3eme n°'
                $ForthNumberLoto = $elementLoto.'4eme n°'
                $FifthNumberLoto = $elementLoto.'5eme n°'
                $CompNumberLoto = $elementLoto.'n° chance 1'
                $ResultArrayLoto = @($FirstNumberLoto,$SecondtNumberLoto,$ThirdNumberLoto,$ForthNumberLoto,$FifthNumberLoto)

                    If(($FirstMyNumberLoto -in $ResultArrayLoto) -and ($SecondtMyNumberLoto -in $ResultArrayLoto) -and ($ThirdMyNumberLoto -in $ResultArrayLoto) -and ($ForthMyNumberLoto -in $ResultArrayLoto) -and ($FifthMyNumberLoto -in $ResultArrayLoto))
                        {

                        if($CompMyNumberLoto -eq $CompNumberLoto)
                            {
                            Write-Host ""
                            Write-Host ""
                            Write-Host "\o/ Niceluuuu ton numéro a été trouvé le $DayLoto $MonthLoto $YearLoto sur les resultats du Loto \o/"
                            Write-Host ""
                            }
                        Else
                            {
                            Write-Host "Presque tu tu t'es trompé sur le numéro complémentaire il s'agissait du combo du $DayLoto $MonthLoto $YearLoto avec le numéro $ResultArrayLoto $CompMyNumberLoto"
                            }
                        }
                    Else
                        {
                        }
                }
        }

        ElseIf ($Source -match "Euromillions")
        {
            $FirstMyNumberEuromillions = $FirstMyNumber
            $SecondtMyNumberEuromillions = $SecondtMyNumber
            $ThirdMyNumberEuromillions = $ThirdMyNumber
            $ForthMyNumberEuromillions = $ForthMyNumber
            $FifthMyNumberEuromillions = $FifthMyNumber
            $CompMyNumberEuromillions1 = $CompMyNumber1
            $CompMyNumberEuromillions2 = $CompMyNumber2
            $MyArrayEuromillions = @($FirstMyNumberEuromillions,$SecondtMyNumberEuromillions,$ThirdMyNumberEuromillions,$ForthMyNumberEuromillions,$FifthMyNumberEuromillions)
            $MyArrayEuromillionsComps = @($CompMyNumberEuromillions1,$CompMyNumberEuromillions2)
            $MyArrayEuromillionsALL = @($FirstMyNumberEuromillions,$SecondtMyNumberEuromillions,$ThirdMyNumberEuromillions,$ForthMyNumberEuromillions,$FifthMyNumberEuromillions,$CompMyNumberEuromillions1,$CompMyNumberEuromillions2)

                Foreach($elementEuromillions in $ResultEuromillionsNumber) 
                {
        
                $DayEuromillions =  $elementEuromillions.Jour
                $MonthEuromillions =  $elementEuromillions.Mois
                $YearEuromillions =  $elementEuromillions.Année
                $FirstNumberEuromillions = $elementEuromillions.'1er n°'
                $SecondtNumberEuromillions = $elementEuromillions.'2eme n°'
                $ThirdNumberEuromillions = $elementEuromillions.'3eme n°'
                $ForthNumberEuromillions = $elementEuromillions.'4eme n°'
                $FifthNumberEuromillions = $elementEuromillions.'5eme n°'
                $CompNumberEuromillions1 = $elementEuromillions.'n° chance 1'
                $CompNumberEuromillions2 = $elementEuromillions.'n° chance 2'
                $ResultArrayEuromillions = @($FirstNumberEuromillions,$SecondtNumberEuromillions,$ThirdNumberEuromillions,$ForthNumberEuromillions,$FifthNumberEuromillions)
                $ResultArrayEuromillionscomps = @($CompNumberEuromillions1,$CompNumberEuromillions2)
        
                    If(($FirstMyNumberEuromillions -in $ResultArrayEuromillions) -and ($SecondtMyNumberEuromillions -in $ResultArrayEuromillions) -and ($ThirdMyNumberEuromillions -in $ResultArrayEuromillions) -and ($ForthMyNumberEuromillions -in $ResultArrayEuromillions) -and ($FifthMyNumberEuromillions -in $ResultArrayEuromillions))
                        {
                        if(($CompMyNumberEuromillions1 -in $ResultArrayEuromillionscomps) -and ($CompMyNumberEuromillions2 -in $ResultArrayEuromillionscomps))
                            {
                            Write-Host "\o/ Niceluuuu ton numéro a été trouvé le $DayEuromillions $MonthEuromillions $YearEuromillions sur les resultats de l'Euromillions \o/"
                            }
                        ElseIf(($CompMyNumberEuromillions1 -inotin $ResultArrayEuromillionscomps) -and ($CompMyNumberEuromillions2 -in $ResultArrayEuromillionscomps))
                            {
                            Write-Host "Presque, mais tu t'es trompé sur le numéro complémentaire 1, il s'agissait du combo du $DayEuromillions $MonthEuromillions $YearEuromillions avec le numéro $ResultArrayEuromillions $ResultArrayEuromillionscomps"
                            }
                        ElseIf(($CompMyNumberEuromillions1 -in $ResultArrayEuromillionscomps) -and ($CompMyNumberEuromillions2 -inotin $ResultArrayEuromillionscomps))
                            {
                            Write-Host "Presque; mais tu t'es trompé sur le numéro complémentaire 2, il s'agissait du combo du $DayEuromillions $MonthEuromillions $YearEuromillions avec le numéro $ResultArrayEuromillions $ResultArrayEuromillionscomps"
                            }
                        Else
                            {
                            Write-Host "Dommage, tu t'es trompé sur les deux numéros complémentaires, il s'agissait du combo du $DayEuromillions $MonthEuromillions $YearEuromillions avec le numéro $ResultArrayEuromillions $ResultArrayEuromillionscomps"
                            }
                        }
                    Else
                        {

                        }
                }
        }
}