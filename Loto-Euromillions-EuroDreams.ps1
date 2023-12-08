#Install-Module -Name ImportExcel
#Install-Module -Name ImportExcel -Scope CurrentUser -Verbose

$ImportExcel = Import-Excel -Path D:\Project\Programs\Powershell\Base.xlsx -WorksheetName "My Numbers"
$ResultLotoNumber = Import-Excel -Path D:\Project\Programs\Powershell\Base.xlsx -WorksheetName "Tirage-Loto"
$ResultEuroMillionsNumber = Import-Excel -Path D:\Project\Programs\Powershell\Base.xlsx -WorksheetName "Tirage-EuroMillions"
$ResultEuroDreamsNumber = Import-Excel -Path D:\Project\Programs\Powershell\Base.xlsx -WorksheetName "Tirage-EuroDreams"


Foreach($values in $ImportExcel) 
{

    $MyArray = @{
        Source=$values.Source;First=$values.'First Number';Second=$values.'Second Number';Third=$values.'Third Number';Forth=$values.'Fourth Number';Fifth=$values.'Fifth Number';sixth=$values.'Sixth Number';Comp1=$values.'n° chance 1';Comp2=$values.'n° chance 2'
    }

    $Source = $values.Source
    $FirstMyNumber = $MyArray.First
    $SecondtMyNumber = $MyArray.Second
    $ThirdMyNumber = $MyArray.Third
    $ForthMyNumber = $MyArray.Forth
    $FifthMyNumber = $MyArray.Fifth
    $SixthMyNumber = $MyArray.Sixth
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

                $DayLoto =  $elementLoto.Day
                $MonthLoto =  $elementLoto.Month
                $YearLoto =  $elementLoto.Year
                $FirstNumberLoto = $elementLoto.'First Number'
                $SecondtNumberLoto = $elementLoto.'Second Number'
                $ThirdNumberLoto = $elementLoto.'Third Number'
                $ForthNumberLoto = $elementLoto.'Fourth Number'
                $FifthNumberLoto = $elementLoto.'Fifth Number'
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
        
                $DayEuromillions =  $elementEuromillions.Day
                $MonthEuromillions =  $elementEuromillions.Month
                $YearEuromillions =  $elementEuromillions.Year
                $FirstNumberEuromillions = $elementEuromillions.'First Number'
                $SecondtNumberEuromillions = $elementEuromillions.'Second Number'
                $ThirdNumberEuromillions = $elementEuromillions.'Third Number'
                $ForthNumberEuromillions = $elementEuromillions.'Fourth Number'
                $FifthNumberEuromillions = $elementEuromillions.'Fifth Number'
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
        ElseIf ($Source -match "EuroDreams")
        {
            $FirstMyNumberEurodreams = $FirstMyNumber
            $SecondtMyNumberEurodreams = $SecondtMyNumber
            $ThirdMyNumberEurodreams = $ThirdMyNumber
            $ForthMyNumberEurodreams = $ForthMyNumber
            $FifthMyNumberEurodreams = $FifthMyNumber
            $SixthMyNumberEurodreams = $SixthMyNumber
            $CompMyNumberEurodreams1 = $CompMyNumber1
            $MyArrayEurodreams = @($FirstMyNumberEurodreams,$SecondtMyNumberEurodreams,$ThirdMyNumberEurodreams,$ForthMyNumberEurodreams,$FifthMyNumberEurodreams,$SixthMyNumberEurodreams)
            $MyArrayEurodreamsComps = @($CompMyNumberEurodreams1)
            $MyArrayEurodreamsALL = @($FirstMyNumberEurodreams,$SecondtMyNumberEurodreams,$ThirdMyNumberEurodreams,$ForthMyNumberEurodreams,$FifthMyNumberEurodreams,$SixthMyNumberEurodreams,$CompMyNumberEurodreams1)

                Foreach($elementEurodreams in $ResultEurodreamsNumber) 
                {
        
                $DayEurodreams =  $elementEurodreams.Day
                $MonthEurodreams =  $elementEurodreams.Month
                $YearEurodreams =  $elementEurodreams.Year
                $FirstNumberEurodreams = $elementEurodreams.'First Number'
                $SecondtNumberEurodreams = $elementEurodreams.'Second Number'
                $ThirdNumberEurodreams = $elementEurodreams.'Third Number'
                $ForthNumberEurodreams = $elementEurodreams.'Fourth Number'
                $FifthNumberEurodreams = $elementEurodreams.'Fifth Number'
                $SixthNumberEurodreams = $elementEurodreams.'Sixth Number'
                $CompNumberEurodreams1 = $elementEurodreams.'n° chance 1'
                $ResultArrayEurodreams = @($FirstNumberEurodreams,$SecondtNumberEurodreams,$ThirdNumberEurodreams,$ForthNumberEurodreams,$FifthNumberEurodreams,$SixthNumberEurodreams)
                $ResultArrayEurodreamscomps = @($CompNumberEurodreams1)
        
                    If(($FirstMyNumberEurodreams -in $ResultArrayEurodreams) -and ($SecondtMyNumberEurodreams -in $ResultArrayEurodreams) -and ($ThirdMyNumberEurodreams -in $ResultArrayEurodreams) -and ($ForthMyNumberEurodreams -in $ResultArrayEurodreams) -and ($FifthMyNumberEurodreams -in $ResultArrayEurodreams) -and ($SixthMyNumberEurodreams -in $ResultArrayEurodreams))
                        {
                        if($CompMyNumberEurodreams1 -in $ResultArrayEurodreamscomps)
                            {
                            Write-Host "\o/ Niceluuuu ton numéro a été trouvé le $DayEurodreams $MonthEurodreams $YearEurodreams sur les resultats de l'Eurodreams \o/"
                            }
                        Else
                            {
                            Write-Host "Presque, mais tu t'es trompé sur le numéro complémentaire , il s'agissait du combo du $DayEurodreams $MonthEurodreams $YearEurodreams avec le numéro $ResultArrayEurodreams $ResultArrayEurodreamscomps"
                            }
                        }
                    Else
                        {

                        }
                }
        }
}