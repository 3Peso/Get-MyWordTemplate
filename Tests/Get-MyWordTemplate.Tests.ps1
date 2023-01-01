Import-Module .\Get-MyWordTemplate.psm1 -Force

InModuleScope Get-MyWordTemplate {
    BeforeAll {
        #$script:verbPref = $VerbosePreference
        #$VerbosePreference = 'Continue'            
    }

    AfterAll {
        #$VerbosePreference = $script:verbPref
    }

    Describe 'Get-MyWordTemplateNames' {   
        Context 'when called with a valid template path' {              
            It 'should return the names of the templates' {           
                Get-MyWordTemplateNames -templatePath ".\Tests\validtemplatedefinitions" | Should -Be 'protocol','testworddoc'
            }
        }
    }

    Describe 'Get-TemplateDefinitionsPath' {  
        Context 'when called with a valid template type' {
            It 'should return the path to the xml file' {     
                # mock Test-MyWordTemplateDefinitionSchema
                Mock Test-MyWordTemplateDefinitionSchema { return $true }
                Mock Get-MatchingFiles { return @(Resolve-Path -path ('.\Tests\validtemplatedefinitions\protocol.xml')) } 
                $result = Get-TemplateDefinitionsPath -templatePath ".\Tests\validtemplatedefinitions" -templateType 'Protocol'            
                $result -like '*\Tests\validtemplatedefinitions\protocol.xml' | Should -Be $true
            }

            It 'should return the path to the xml file even if the inner function Get-MatchingFiles only returns a single string.' {
                # mock Test-MyWordTemplateDefinitionSchema
                Mock Test-MyWordTemplateDefinitionSchema { return $true }
                Mock Get-MatchingFiles { return '.\Tests\validtemplatedefinitions\protocol.xml' } 
                Mock Test-Path { return $true }
                $result = Get-TemplateDefinitionsPath -templatePath ".\Tests\validtemplatedefinitions" -templateType 'Protocol'            
                $result | Should -Be '.\Tests\validtemplatedefinitions\protocol.xml'
            }
        }
        Context 'when called with an invalid template type' {
            It 'should throw an exception' { 
                # mock Test-MyWordTemplateDefinitionSchema
                Mock Test-MyWordTemplateDefinitionSchema { return $true }      
                Mock Get-MatchingFiles { return @() }                         
                {Get-TemplateDefinitionsPath -templatePath ".\Tests\invalidtemplatedefinitions" -templateType 'invalid'} | Should -Throw 'No file matches the template type invalid'
            }
        }
        Context 'when called with a template type that matches more than one file' {
            It 'should throw an exception' {     
                # mock Test-MyWordTemplateDefinitionSchema
                Mock Test-MyWordTemplateDefinitionSchema { return $true }   
                Mock Get-MatchingFiles { return @('File1.xml', 'File2.xml') }                     
                {Get-TemplateDefinitionsPath -templatePath ".\Tests\invalidtemplatedefinitions" -templateType 'template2'} | Should -Throw 'More than one file matches the template type template2'
            }
        }
    }

    Describe 'Get-TemplateInput' {  
        Context 'when called with a valid xml' {
            It 'should return the template input with one input' {     
                # Mock Invoke-TemplateConfigElement
                Mock Invoke-TemplateConfigElement { return @{'Hello' = 'World'}}
                $xml = [xml]@"
                <MyWordTemplate $script:TEMPLATE_DEFINITION_NAME="TestTemplate">
                <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="SomeDate" $script:USER_PROMPT="Enter Date" $script:ELEMENT_ID="SomeDate"/>
                <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="SomeData" $script:USER_PROMPT="Enter Number" $script:ELEMENT_ID="SomeData"/>
            </MyWordTemplate>
"@
                $result = Get-TemplateInput -xml $xml.DocumentElement
                $result.Count | Should -Be 1
            }

            It 'should return the template input with two inputs' {     
                # mock Read-Host
                Mock Read-Host { return 'Enter a value' }
                $xml = [xml]@"
                <MyWordTemplate $script:TEMPLATE_DEFINITION_NAME="TestTemplate">
                <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="SomeDate" $script:USER_PROMPT="Enter Date" $script:ELEMENT_ID="SomeDate" />
                <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="SomeData" $script:USER_PROMPT="Enter Number" $script:ELEMENT_ID="SomeData" />
            </MyWordTemplate>
"@
                $result = Get-TemplateInput -xml $xml.DocumentElement        
                $result.Count | Should -Be 2
            }

            It 'should return the template input with two inputs and the Placeholder elements inner text' {
                # mock Read-Host
                Mock Read-Host { return 'Enter a value' }
                $xml = [xml]@"
                <MyWordTemplate $script:TEMPLATE_DEFINITION_NAME="TestTemplate">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="SomeDate" $script:USER_PROMPT="Enter Date" $script:ELEMENT_ID="SomeDate" />
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="SomeData" $script:USER_PROMPT="Enter Number" $script:ELEMENT_ID="SomeData" />
                    <Placeholder $script:ELEMENT_ID="SomePlaceholder">SomeData</Placeholder>
                </MyWordTemplate>
"@
                $result = Get-TemplateInput -xml $xml.DocumentElement        
                $result.Count | Should -Be 3
            }       

            It 'should return the template input with a valid input if attribute ValidateRegex was set' {     
                # mock Read-Host
                # should return user input "1Hello" which is a valid input
                Mock Read-Host { return '1' }
                Mock Test-HasCondition { return $false }
                $xml = [xml]@"
                <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="SomeDate" $script:USER_PROMPT="Enter Date" $script:VALIDATION_REGEX="^[0-9]+$" $script:ELEMENT_ID="SomeData" />
"@
                $result = Get-TemplateInput -xml $xml.DocumentElement        
                $result.Count | Should -Be 1
                $result.Values[0] | Should -Be '1'
            }

            AfterEach {                 
                $script:WordTemplateInput = @{}
            }
        }
    }

    Describe 'Get-UserInput' {  
        Context 'when called with a valid xml' {
            It 'should return the user input' {  
                # mock Read-Host
                Mock Read-Host { return 'Enter a value' } 
                # mock Test-UserInputCondition
                Mock Test-UserInputCondition { return $true }  
                [Xml]$xml = @"
                <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test" $script:USER_PROMPT="Enter a value" $script:ELEMENT_ID="SomeData"/>
"@
                $result = Get-UserInput -inputElement $xml.DocumentElement
                $result.Count | Should -Be 1
                $result.SomeData | Should -Be 'Enter a value'
            }

            It 'should validate the user input with a provided regex' {
                # mock Read-Host
                Mock Read-Host { return '123' }   
                Mock Test-UserInputCondition { return $true }                  
                [Xml]$xml = @"
                <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test" $script:USER_PROMPT="Enter a value" $script:VALIDATION_REGEX="^[0-9]+$" $script:ELEMENT_ID="SomeData"/>
"@
                $result = Get-UserInput -inputElement $xml.DocumentElement
                $result.Count | Should -Be 1
                $result.SomeData | Should -Be '123'
            }
        }
    }

    Describe 'Test-MyWordTemplateDefinitionSchema' {  
        Context 'when called with a valid xml' {
            It 'should return true' {  
                [string]$xmlFile = ".\Tests\validtemplatedefinitions\protocol.xml"
                $result = Test-MyWordTemplateDefinitionSchema -xmlFilePath $xmlFile
                $result | Should -Be $true
            }                      
        }

        Context 'when called with an invalid xml' {
            It 'should return false' {  
                [string]$xmlFile = ".\Tests\invalidtemplatedefinitions\invalidTemplateDefinition.xml"
                $result = Test-MyWordTemplateDefinitionSchema -xmlFilePath $xmlFile
                $result | Should -Be $false
            }
        }
    }

    Describe 'Invoke-TemplateConfigElement' {  
        Context 'when called with a valid xml' {
            It 'should add the placeholder elements inner text to the template input' {     
                $xml = [xml]@"
                <Placeholder $script:TEMPLATE_DEFINITION_NAME="SomePlaceholder" $script:ELEMENT_ID="SomePlaceholder">Some Placeholder Text</Placeholder>
"@
                $result = Invoke-TemplateConfigElement -templateElement $xml.DocumentElement        
                $result.Count | Should -Be 1
                $result.Values[0] | Should -Be 'Some Placeholder Text'
            }

            It 'should return a hashtable with the id of a loop element as key and a nested hashtable as value' {     
                $xml = [xml]@"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="SomeLoop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="SomeDate" />                                  
                </$script:LOOP_INPUT_ELEMENT>
"@
                # mock Get-LoopInput
                Mock Get-LoopInput{ return [ordered]@{ 'SomeDate' = 'Test'
                                                       'AnotherDate' = 'Test2' } }
                $result = Invoke-TemplateConfigElement -templateElement $xml.DocumentElement        
                $result.Count | Should -Be 1
                $result.Keys[0] | Should -Be 'SomeLoop'
                $result.Values[0].Keys.Count | Should -Be 2
                $result.Values[0].Values[0] | Should -Be 'Test'
                $result.Values[0].Values[1] | Should -Be 'Test2'                
            }

            It 'should return nothing if the ID of the element is already a key in $script:WordTemplateInput' {     
                $xml = [xml]@"
                <Placeholder $script:TEMPLATE_DEFINITION_NAME="SomePlaceholder" $script:ELEMENT_ID="SomePlaceholder">Some Placeholder Text</Placeholder>
"@
                $script:WordTemplateInput = [ordered]@{ 'SomePlaceholder' = 'Some Placeholder Text' }
                $result = Invoke-TemplateConfigElement -templateElement $xml.DocumentElement        
                $result.Count | Should -Be 0
            }
        }
    }

    Describe 'Test-MyWordTemplateDefinitions' {  
        Context 'when called with a valid xml' {
            It 'should return an array with 0 entries' {     
                Mock Test-TemplateDefinitionFilename { return $true } 
                $result = Test-MyWordTemplateDefinitions -templateDefinitionPath '.\Tests\validtemplatedefinitions'    
                $result.Count | Should -Be 0
            }
        }
56
        Context 'when called with an invalid xml' {
            It 'should return an array with three entries' {   
                Mock Test-TemplateDefinitionFilename { return $true }  
                $result = Test-MyWordTemplateDefinitions -templateDefinitionPath ".\Tests\invalidtemplatedefinitions"     
                $result.Count | Should -Be 1
            }
        }
    }

    Describe 'Test-WordTemplatesAgainstWordTemplateDefintions' {  
        Context 'when every template has matching template definition' {
            It 'should return true' {     
                $result = Test-WordTemplatesAgainstWordTemplateDefintions -templateDefinitionsPath '.\Tests\validtemplatedefinitions' -templatesPath '.\tests\validtemplates'    
                $result | Should -Be $true
            }
        }

        Context 'when called a template is missing' {
            It 'should return false' {    
                # Mock Get-FileNamesInFolder if parameter filter is *.xml
                Mock Get-FileNamesInFolder { return @('TestTemplate.xml') } -ParameterFilter { $filter -eq '*.xml' }
                Mock Test-TemplateDefinitionFilename { return $true }  
                Write-Host "Testing for missing template ..." -ForegroundColor Yellow                   
                $result = Test-WordTemplatesAgainstWordTemplateDefintions -templateDefinitionsPath ".\Tests\validtemplatedefinitions" -templatesPath ".\Tests\validtemplates"     
                $result | Should -Be $false
            }
        }      
    }

    Describe 'Test-MyWordTemplateInstallation' {
        Context 'when installation is not valid' {
            It 'should throw exception if word is not installed' {
                # Mock Test-WordInstallation
                Mock Test-WordInstallation { return $false }
                Mock Test-MyWordTemplateDefinitions { return @() } 
                Mock Test-WordTemplatesAgainstWordTemplateDefintions { return $true }         
                {Test-MyWordTemplateInstallation} | Should -Throw 'Word is not installed.'
            }

            It 'should throw exception if word template definitions are not valid' {
                # Mock Test-WordInstallation
                Mock Test-WordInstallation { return $true }
                Mock Test-MyWordTemplateDefinitions { return @('Some error') }
                Mock Test-WordTemplatesAgainstWordTemplateDefintions { return $true }                  
                {Test-MyWordTemplateInstallation} | Should -Throw 'One or more word template definitions are not valid.'
            }

            It 'should throw exception if word template is missing' {
                # Mock Test-WordInstallation
                Mock Test-WordInstallation { return $true }
                Mock Test-MyWordTemplateDefinitions { return @() }
                Mock Test-WordTemplatesAgainstWordTemplateDefintions { return $false }                  
                {Test-MyWordTemplateInstallation} | Should -Throw 'One or more word templates are missing.'
            }
        }

        Context 'when installation is valid' {
            It 'should return true if word is installed' {
                # Mock Test-WordInstallation
                Mock Test-WordInstallation { return $true }
                Mock Test-MyWordTemplateDefinitions { return @() }
                Mock Test-WordTemplatesAgainstWordTemplateDefintions { return $true }
                $result = Test-MyWordTemplateInstallation
                $result | Should -Be $true
            }
        }
    }

    Describe 'Get-DocumentOutputPath' {
        Context 'when called with a file and path which already exists' {
            It 'should return an output path with a counter of 1 at the end of the file name' {     
                # mock Test-Path
                Mock Test-Path { return $true }
                Mock Test-Path { return $false } -ParameterFilter { $path -eq '.\GeneratedDocuments\protocol(1).docx' }
                Mock Resolve-Path { return '.\GeneratedDocuments'}
                $result = Get-DocumentOutputPath -templatePath ".\templtes\validtemplatedfinitions\protocol.xml" -outputPath ".\GeneratedDocuments"
                $result | Should -Be '.\GeneratedDocuments\protocol(1).docx'
            }
        }

        Context 'when called with a file and path which does not exist' {
            It 'should return the path to the word document' {     
                # mock Test-Path
                Mock Test-Path { return $false } -ParameterFilter { $path -eq '.\GeneratedDocuments\protocol.docx' }
                Mock Test-Path { return $true }
                Mock Resolve-Path { return '.\GeneratedDocuments'}                
                $result = Get-DocumentOutputPath -templatePath ".\templtes\validtemplatedfinitions\protocol.xml" -outputPath ".\GeneratedDocuments"
                $result | Should -Be '.\GeneratedDocuments\protocol.docx'
            }
        }
    }

    Describe 'New-MyWordDocument' {
        Context 'when called with a valid template' {
            It 'should create a word document' {     
                # mock Test-Path
                Mock Test-Path { return $false } -ParameterFilter { $path -eq '.\Tests\GeneratedDocuments\testworddoc.docx' }
                Mock Test-Path { return $true }
                Mock Get-DocumentOutputPath { return Join-Path $(Resolve-Path -path '.\Tests\GeneratedDocuments') 'testworddoc.docx' }
                Mock Test-MyWordTemplateDefinitionSchema { return $true }
                $loopInput = [ordered]@{loopentry1 = 'entry 1'
                                        loopentry2 = 'entry 2'}
                $templateInput = @{test = 'TEST'
                                   testloop = $loopInput
                                   test2 = 'TEST2'}
                $result = New-MyWordDocument -templatePath ".\Tests\validtemplatedefinitions\testworddoc.xml" -wordTemplatePath ".\Tests\validtemplates\testworddoc.docx" -outputPath ".\Tests\GeneratedDocuments" -templateInput $templateInput
                $result | Should -Be $true
            }

            AfterEach {
                $testfile = Resolve-Path -path '.\Tests\GeneratedDocuments\testworddoc.docx' -ErrorAction SilentlyContinue
                if (Test-Path -Path $testfile) {
                    Remove-Item -Path $testfile -Force -ErrorAction SilentlyContinue
                } else {
                    Write-Host "File '$testfile' does not exist." -ForegroundColor Yellow
                }
            }
        }
    }

    Describe 'Test-HasCondition' {
        Context 'when called with a valid condition' {
            It 'should return true' {                 
                [xml]$inputElement = @"
                <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Auftragseigang" $script:USER_PROMPT="Datum des Auftragseingangs (z.B. 01.01.2019)" $script:ELEMENT_ID="auftragseingang" $script:VALIDATION_REGEX="^[0-9]{2}.[0-9]{2}.[0-9]{4}$" $script:YOUNGER_DATE_CONDITION="auftragsdatum" />
"@  
                $result = Test-HasCondition -xmlElement $inputElement.DocumentElement
                $result | Should -Be $true
            }
        }

        Context 'when called with an invalid condition' {
            It 'should return false' {     
                [xml]$inputElement = @"
                <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Auftragseigang" $script:USER_PROMPT="Datum des Auftragseingangs (z.B. 01.01.2019)" $script:ELEMENT_ID="auftragseingang" $script:VALIDATION_REGEX="^[0-9]{2}.[0-9]{2}.[0-9]{4}$" />
"@ 
                $result = Test-HasCondition -xmlElement $inputElement.DocumentElement
                $result | Should -Be $false
            }
        }
    }

    Describe 'Test-InputDate' {
        Context 'when called with a valid date' {
            It 'should return true' {     
                [DateTime]$auftragsdatum = '12.12.2018'
                $result = Test-InputDate -dateTable @{auftragsdatum=$auftragsdatum} -dateString '01.01.2019' -dateTableKey 'auftragsdatum'
                $result | Should -Be $true
            }

            It 'should return true if dates are the same' {     
                [DateTime]$auftragsdatum = '01.01.2019'
                $result = Test-InputDate -dateTable @{auftragsdatum=$auftragsdatum} -dateString '01.01.2019' -dateTableKey 'auftragsdatum'
                $result | Should -Be $true
            }            
        }

        Context 'when called with an invalid date' {
            It 'should return false' {     
                [DateTime]$auftragsdatum = '02.01.2019' 
                $result = Test-InputDate -dateTable @{auftragsdatum=$auftragsdatum} -dateString '01.01.2019' -dateTableKey 'auftragsdatum'
                $result | Should -Be $false
            }
        }
    }

    Describe 'Test-UserInputCondition' {
        Context 'when called with a valid condition' {
            It 'should return $true' {   
                [xml]$inputElement = @"
                <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Auftragseigang" $script:USER_PROMPT="Datum des Auftragseingangs (z.B. 01.01.2019)" $script:ELEMENT_ID="auftragseingang" $script:VALIDATION_REGEX="^[0-9]{2}.[0-9]{2}.[0-9]{4}$" $script:YOUNGER_DATE_CONDITION="auftragsdatum" />
"@  
                $script:WordTemplateInput = @{auftragsdatum='20.12.2018'}
                $result = Test-UserInputCondition -inputString '01.01.2019' -inputElement $inputElement.DocumentElement
                $result | Should -Be $true
                $script:WordTemplateInput.auftragsdatum | Should -Be '20.12.2018'
            }
        }

        Context 'when called with an invalid condition' {
            It 'should be false' {
                # mock Test-InputDate
                Mock Test-InputDate { return $false }
                [xml]$inputElement = @"
                <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Auftragseigang" $script:USER_PROMPT="Datum des Auftragseingangs (z.B. 01.01.2019)" $script:ELEMENT_ID="auftragseingang" $script:VALIDATION_REGEX="^[0-9]{2}.[0-9]{2}.[0-9]{4}$" $script:YOUNGER_DATE_CONDITION="auftragsdatum" />
"@
                $script:WordTemplateInput = @{auftragsdatum='02.01.2019'}                
                Test-UserInputCondition -inputString '01.01.2019' -inputElement $inputElement.DocumentElement | Should -Be $false
                $script:WordTemplateInput.auftragsdatum | Should -Be '02.01.2019'
            }

            AfterEach {
                $script:WordTemplateInput = @{}
            }
        }
    }

    Describe 'Get-WordTemplatePath' {
        Context 'when called with a valid template' {
            It 'should return the path to the template' {     
                Mock Test-Path { return $true }
                $result = Get-WordTemplatePath -templateDefinitionFile ".\Tests\validtemplatedefinitions\testworddoc.xml" -wordTemplatePath ".\Tests\validtemplates"
                $result | Should -BeLike "*\Tests\validtemplates\testworddoc.docx"
            }
        }

        Context 'when called with an invalid template' {
            It 'should throw an error' {     
                Mock Test-Path { return $true }
                Mock Get-ChildItem { return ".\Tests\testtemplates\testword.docx" }                
                { Get-WordTemplatePath -templateDefinitionFile ".\Tests\testworddoc.xml" -wordTemplatePath ".\Tests\testtemplates" } | Should -Throw
            }
        }
    }

    Describe 'Get-LoopInput and Get-LoopChild' {
        Context 'when called with a valid loop' {
            It 'should return a hashtable with two entries' {
                [xml]$inputElement = @"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />  
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry ID" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentryid" />                                    
                </$script:LOOP_INPUT_ELEMENT>
"@
                $script:mockCounter = 0
                Mock Invoke-TemplateConfigElement {
                    $returnValue = $null
                    if ($mockCounter -eq 0) {
                        $returnValue = @{testentry='test'}
                    }
                    if ($mockCounter -eq 1) {
                        $returnValue = @{testentryid='testid'}
                    }
                    if ($mockCounter -eq 2) {
                        $returnValue = @{testentry='done'}
                    }
                    $script:mockCounter++
                    return $returnValue
                }

                $result = Get-LoopInput -inputElement $inputElement.DocumentElement
                $result.Count | Should -Be 3
                $result.Keys[2] | Should -be "1$script:LOOPEND_MARKER"
            }

            It 'should return a hashtable with three entries' {
                [xml]$inputElement = @"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />  
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry ID" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentryid" />                                    
                </$script:LOOP_INPUT_ELEMENT>
"@
                $script:mockCounter = 0
                Mock Invoke-TemplateConfigElement {
                    $returnValue = $null
                    if ($mockCounter -eq 0) {
                        $returnValue = @{testentry='test'}
                    }
                    if ($mockCounter -eq 1) {
                        $returnValue = @{testentryid='testid'}
                    }
                    if ($mockCounter -eq 2) {
                        $returnValue = @{testentry='test'}
                    }                    
                    if ($mockCounter -eq 3) {
                        $returnValue = @{testentry='done'}
                    }
                    $script:mockCounter++
                    return $returnValue
                }

                $result = Get-LoopInput -inputElement $inputElement.DocumentElement
                $result.Count | Should -Be 4
                $result.Keys[2] | Should -Be "1$script:LOOPEND_MARKER"
                $result.Keys[3] | Should -Be '2testentry'
            } 
            
            # REMARKS: This test is important, because if it is not met, it will break the function Invoke-TemplateConfigElement
            It 'should return an OrderdDictionary, not a hashtable' {
                [xml]$inputElement = @"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />  
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry ID" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentryid" />                                    
                </$script:LOOP_INPUT_ELEMENT>
"@
                $script:mockCounter = 0
                Mock Invoke-TemplateConfigElement {
                    $returnValue = $null
                    if ($mockCounter -eq 0) {
                        $returnValue = @{testentry='test'}
                    }
                    if ($mockCounter -eq 1) {
                        $returnValue = @{testentryid='testid'}
                    }
                    if ($mockCounter -eq 2) {
                        $returnValue = @{testentry='done'}
                    }
                    $script:mockCounter++
                    return $returnValue
                }

                $result = Get-LoopInput -inputElement $inputElement.DocumentElement
                $result.GetType().Name | Should -Be 'OrderedDictionary'
            }
        }

        Context 'when called with an invalid loop' {
            It 'should throw an error if the id attributes in the inner elements are the same' {
                [xml]$inputElement = @"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />  
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry ID" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />                                    
                </$script:LOOP_INPUT_ELEMENT>
"@
                Mock Invoke-TemplateConfigElement { return @{testentry='test'} }
                { Get-LoopInput-loopElement $inputElement.DocumentElement } | Should -Throw
            }

            It 'should throw an error if there are two levels or more of elements inside the loop' {
                [xml]$inputElement = @"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />
                    <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                        <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />  
                    </$script:LOOP_INPUT_ELEMENT>
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry ID" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />                                    
                </$script:LOOP_INPUT_ELEMENT>
"@
                Mock Invoke-TemplateConfigElement { return @{testentry='test'} }
                { Get-LoopInput-loopElement $inputElement.DocumentElement } | Should -Throw
            }  
            
            It 'should throw an error if the input does not return a result' {
                [xml]$inputElement = @"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />                               
                </$script:LOOP_INPUT_ELEMENT>
"@
                Mock Invoke-TemplateConfigElement { return @{} }
                { Get-LoopInput-loopElement $inputElement.DocumentElement } | Should -Throw
            }               
        }
    }

    Describe 'Get-LoopChild' {
        Context 'when called with a valid loop' {
            It 'should return a hashtable with two entries' {
                [xml]$loopElement = @"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />  
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry ID" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentryid" />                                    
                </$script:LOOP_INPUT_ELEMENT>
"@
                $script:mockCounter = 0
                Mock Invoke-TemplateConfigElement {
                    $returnValue = $null
                    if ($mockCounter -eq 0) {
                        $returnValue = @{testentry='test'}
                    }
                    if ($mockCounter -eq 1) {
                        $returnValue = @{testentryid='testid'}
                    }
                    if ($mockCounter -eq 2) {
                        $returnValue = @{testentry='done'}
                    }
                    $script:mockCounter++
                    return $returnValue
                }

                $result = Get-LoopChild -loopElement $loopElement.DocumentElement -loopElementId 'testloop' -breakkeyword 'done' -loopCounter 1
                $result.Count | Should -Be 2
            }

            It 'should return a hashtable with three entries' {
                [xml]$inputElement = @"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />  
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry ID" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentryid" />                                    
                </$script:LOOP_INPUT_ELEMENT>
"@
                $script:mockCounter = 0
                Mock Invoke-TemplateConfigElement {
                    $returnValue = $null
                    if ($mockCounter -eq 0) {
                        $returnValue = @{testentry='test'}
                    }
                    if ($mockCounter -eq 1) {
                        $returnValue = @{testentryid='testid'}
                    }
                    if ($mockCounter -eq 2) {
                        $returnValue = @{testentry='test'}
                    }                    
                    if ($mockCounter -eq 3) {
                        $returnValue = @{testentry='done'}
                    }
                    $script:mockCounter++
                    return $returnValue
                }
                
                $result = Get-LoopChild -loopElement $inputElement.DocumentElement -loopElementId 'testloop' -breakkeyword 'done' -loopCounter 2
                Assert-MockCalled Invoke-TemplateConfigElement -Times 2 -Exactly -Scope It
                # This must be 2 not 3 because Get-LoopChild will only loop over the child elements of the loop element once
                $result.Count | Should -Be 2
                $result.Keys[1] | Should -Be '2testentryid'
            }           
        }

        Context 'when called with an invalid loop' {
            It 'should throw an error if the id attributes in the inner elements are the same' {
                [xml]$inputElement = @"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />  
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry ID" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />                                    
                </$script:LOOP_INPUT_ELEMENT>
"@
                Mock Invoke-TemplateConfigElement { return @{testentry='test'} }
                { Get-LoopChild -loopElement $inputElement.DocumentElement -loopElementId 'testloop' -breakkeyword 'done' -loopCounter 1 } | Should -Throw
            }

            It 'should throw an error if there are two levels or more of elements inside the loop' {
                [xml]$inputElement = @"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />
                    <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                        <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />  
                    </$script:LOOP_INPUT_ELEMENT>
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry ID" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />                                    
                </$script:LOOP_INPUT_ELEMENT>
"@
                Mock Invoke-TemplateConfigElement { return @{testentry='test'} }
                { Get-LoopChild -loopElement $inputElement.DocumentElement -loopElementId 'testloop' -breakkeyword 'done' -loopCounter 1 } | Should -Throw
            }  
            
            It 'should throw an error if the input does not return a result' {
                [xml]$inputElement = @"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Loop" $script:ELEMENT_ID="testloop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="Loop Prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="Test Entry" $script:USER_PROMPT="Test Entry" $script:ELEMENT_ID="testentry" />                               
                </$script:LOOP_INPUT_ELEMENT>
"@
                Mock Invoke-TemplateConfigElement { return @{} }
                { Get-LoopChild -loopElement $inputElement.DocumentElement -loopElementId 'testloop' -breakkeyword 'done' -loopCounter 1 } | Should -Throw
            }                 
        }
    }

    Describe 'Get-MyWordTemplate' {
        Context 'when called with more than one template type, testing the different ways of handling the pipeline variable $templateType' {
            It 'should create a new document for every template type with two template types' {
                #Mock Get-MyWordTemplateNames { return @('Document1','Document2') }
                Mock Get-TemplateDefinitionsPath { return '.\Tests\validtemplatedefinitions' }
                Mock Get-Content { return @('<xml></xml>') }
                Mock Get-TemplateInput { return @{test='test'} }
                Mock Get-WordTemplatePath { return '.\Tests\validtemplates\protocol.docx' }
                Mock New-MyWordDocument { return $null }
                Mock Add-TemplateInput { return $true }
                
                Get-MyWordTemplate -templateType 'Document1','Document2'

                Assert-MockCalled Get-TemplateDefinitionsPath -Exactly -Times 1 -Scope It -ParameterFilter { $templateType -eq 'Document1' }
                Assert-MockCalled Get-TemplateDefinitionsPath -Exactly -Times 1 -Scope It -ParameterFilter { $templateType -eq 'Document2' }
                Assert-MockCalled Get-Content -Exactly -Times 2 -Scope It
                Assert-MockCalled Get-TemplateInput -Exactly -Times 2 -Scope It
                Assert-MockCalled Get-WordTemplatePath -Exactly -Times 2 -Scope It
                Assert-MockCalled New-MyWordDocument -Exactly -Times 2 -Scope It
                Assert-MockCalled Add-TemplateInput -Exactly -Times 2 -Scope It
            }

            It 'should create a new document for every template type with two template types as array' {
                #Mock Get-MyWordTemplateNames { return @('Document1','Document2') }
                Mock Get-TemplateDefinitionsPath { return '.\Tests\validtemplatedefinitions' }
                Mock Get-Content { return @('<xml></xml>') }
                Mock Get-TemplateInput { return @{test='test'} }
                Mock Get-WordTemplatePath { return '.\Tests\validtemplates\protocol.docx' }
                Mock New-MyWordDocument { return $null }
                Mock Add-TemplateInput { return $true }              
                
                Get-MyWordTemplate -templateType @('Document1','Document2')

                Assert-MockCalled Get-TemplateDefinitionsPath -Exactly -Times 1 -Scope It -ParameterFilter { $templateType -eq 'Document1' }
                Assert-MockCalled Get-TemplateDefinitionsPath -Exactly -Times 1 -Scope It -ParameterFilter { $templateType -eq 'Document2' }
                Assert-MockCalled Get-Content -Exactly -Times 2 -Scope It
                Assert-MockCalled Get-TemplateInput -Exactly -Times 2 -Scope It
                Assert-MockCalled Get-WordTemplatePath -Exactly -Times 2 -Scope It
                Assert-MockCalled New-MyWordDocument -Exactly -Times 2 -Scope It
                Assert-MockCalled Add-TemplateInput -Exactly -Times 2 -Scope It
            }

            It 'should create a new document for every template type with one template' {
                #Mock Get-MyWordTemplateNames { return @('Document1') }
                Mock Get-TemplateDefinitionsPath { return '.\Tests\validtemplatedefinitions' }
                Mock Get-Content { return @('<xml></xml>') }
                Mock Get-TemplateInput { return @{test='test'} }
                Mock Get-WordTemplatePath { return '.\Tests\validtemplates\protocol.docx' }
                Mock New-MyWordDocument { return $null }
                Mock Add-TemplateInput { return $true }               
                
                Get-MyWordTemplate -templateType 'Document1'

                Assert-MockCalled Get-TemplateDefinitionsPath -Exactly -Times 1 -Scope It -ParameterFilter { $templateType -eq 'Document1' }
                Assert-MockCalled Get-Content -Exactly -Times 1 -Scope It
                Assert-MockCalled Get-TemplateInput -Exactly -Times 1 -Scope It
                Assert-MockCalled Get-WordTemplatePath -Exactly -Times 1 -Scope It
                Assert-MockCalled New-MyWordDocument -Exactly -Times 1 -Scope It
                Assert-MockCalled Add-TemplateInput -Exactly -Times 1 -Scope It
            } 
            
            It 'should create a new document for every template type with two template types passed in by pipeline' {
                #Mock Get-MyWordTemplateNames { return  @('Document1','Document2')}
                Mock Get-TemplateDefinitionsPath { return '.\Tests\validtemplatedefinitions' }
                Mock Get-Content { return @('<xml></xml>') }
                Mock Get-TemplateInput { return @{test='test'} }
                Mock Get-WordTemplatePath { return '.\Tests\validtemplates\protocol.docx' }
                Mock New-MyWordDocument { return $null }
                Mock Add-TemplateInput { return $true }             
                
                @('Document1','Document2') | Get-MyWordTemplate

                Assert-MockCalled Get-TemplateDefinitionsPath -Exactly -Times 2 -Scope It
                Assert-MockCalled Get-Content -Exactly -Times 2 -Scope It
                Assert-MockCalled Get-TemplateInput -Exactly -Times 2 -Scope It
                Assert-MockCalled Get-WordTemplatePath -Exactly -Times 2 -Scope It
                Assert-MockCalled New-MyWordDocument -Exactly -Times 2 -Scope It
                Assert-MockCalled Add-TemplateInput -Exactly -Times 2 -Scope It
            }  
        }

        Context 'End to end tests' {
            It 'should generate two word documents in .\Tests\GeneratedDocuments' {
                Mock Read-Host { return 'test' }
                Get-MyWordTemplate -templateType 'protocol','testworddoc' -outputpath '.\Tests\GeneratedDocuments' -wordTemplatePath '.\Tests\validtemplates' -templatePath '.\Tests\validtemplatedefinitions'
            }
            
            AfterEach {
                $generatedFiles = Get-ChildItem -Path '.\Tests\GeneratedDocuments'
                $generatedFiles | ForEach-Object { Write-Verbose "File $($_.Name) was generated." }
                Write-Verbose "Removing generated files."
                $generatedFiles | Remove-Item
            }
        }
    }

    Describe 'Test-TemplateDefinitionFilename' {
        Context 'when called with a valid template definition filename' {
            It 'should return true' {
                Test-TemplateDefinitionFilename -templateDefinitionFilePath '.\Tests\validtemplatedefinitions\protocol.xml' | Should -BeTrue
            }
        }

        Context 'when called with an invalid template definition filename' {
            It 'should return false' {
                Test-TemplateDefinitionFilename -templateDefinitionFilePath '.\Tests\invalidtemplatedefinitions\invalidtemplatedefinition_byname.xml' | Should -BeFalse
            }
        }
    }

    Describe 'Add-TemplateInput' {
        Context 'when called with user input that is not yet in $script:WordTemplateInput' {
            It 'should add the user input to $script:WordTemplateInput' {
                $script:WordTemplateInput = @{test='test'}
                $result = Add-TemplateInput -userInput @{test2='test2'}
                $result | Should -BeTrue
                $script:WordTemplateInput['test'] | Should -Be 'test'
                $script:WordTemplateInput['test2'] | Should -Be 'test2'
            }

            It 'should overwrite the user input in $script:WordTemplateInput if the user enters "y"' {
                $script:WordTemplateInput = @{test='test'}
                Mock Read-Host { return 'y' }
                $result = Add-TemplateInput -userInput @{test='test2'}
                $result | Should -BeTrue
                $script:WordTemplateInput['test'] | Should -Be 'test2'
            }

            It 'should not overwrite the user input in $script:WordTemplateInput if the user enters "n"' {
                $script:WordTemplateInput = @{test='test'}
                Mock Read-Host { return 'n' }
                $result = Add-TemplateInput -userInput @{test='test2'}
                $result | Should -BeTrue
                $script:WordTemplateInput['test'] | Should -Be 'test'
            }     
            
            It 'should return false and do nothing if the user enters "c"' {
                $script:WordTemplateInput = @{test='test'}
                Mock Read-Host { return 'c' }
                $result = Add-TemplateInput -userInput @{test='test2'}
                $result | Should -BeFalse
                $script:WordTemplateInput['test'] | Should -Be 'test'
            }          

            It 'should throw an exception if the user enters anything other than "y", "n", or "c"' {
                $script:WordTemplateInput = @{test='test'}
                Mock Read-Host { return 'x' }
                { Add-TemplateInput -userInput @{test='test2'} } 4>&1 | Should -Throw
            }

            It 'should throw an exception if the user enters anything other than "y", "n", or "c". $script:WordTemplateInput should be unchanged' {
                $script:WordTemplateInput = [ordered]@{test='test';test2='testX';test3='testX'}
                Mock Read-Host { return 'x' }
                { Add-TemplateInput -userInput @{test2='test2';test3='test2'} } | Should -Throw
                $script:WordTemplateInput['test'] | Should -Be 'test'
                $script:WordTemplateInput['test2'] | Should -Be 'testX'
                $script:WordTemplateInput['test3'] | Should -Be 'testX'
            }  

            AfterEach {
                $script:WordTemplateInput = $null
            }
        }
    }

    Describe 'Get-TemplateInputRecursive' {
        Context 'when called with LoopInput xml element' {
            It 'should return a hashtable with the user input' {
                $xml = [xml]@"
                <$script:LOOP_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="loop" $script:ELEMENT_ID="loop" $script:USER_LOOP_BREAK_SIGNAL="done" $script:USER_PROMPT="loop prompt">
                    <$script:USER_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="user input" $script:USER_PROMPT="user input" $script:ELEMENT_ID="userinput" />
                </$script:LOOP_INPUT_ELEMENT>
"@
                # mock Invoke-TemplateConfigElement
                Mock Invoke-TemplateConfigElement { return @{} }
                $result = Get-TemplateInputRecursive -xml $xml.DocumentElement
                $result.Keys.Count | Should -Be 0
            }
        }
    }

    Describe 'Get-ChoiceInput' {
        Context 'when called with a valid choice' {
            It 'should return the choice' {
                [xml]$xml = "<xml />"
                Mock Build-ChoiceTable { return @{'1'='choice1';'2'='choice2'} }
                Mock Get-UserChoices { return @{choice='choice1'} }
                Mock Read-Host { return '1' }
                $result = Get-ChoiceInput -inputElement $xml.DocumentElement
                $result["choice"] | Should -Be "choice1"
            }

            It 'should return the choice' {
                [xml]$xml = "<xml />"
                Mock Build-ChoiceTable { return @{'1'='choice1';'2'='choice2';'3'='choice3'} }
                Mock Get-UserChoices { return @{choice='choice2'} }                
                Mock Read-Host { return '2' }
                $result = Get-ChoiceInput -inputElement $xml.DocumentElement
                $result["choice"] | Should -Be "choice2"
            }            

            It 'should throw an exception if user enters break signal' {
                [xml]$xml = "<xml />"
                Mock Get-UserChoices { throw "cancel entered" }
                Mock Build-ChoiceTable { return @{'1'='choice1';'2'='choice2';'3'='choice3'} }
                Mock Read-Host { return 'cancel' }
                { Get-ChoiceInput -inputElement $xml.DocumentElement } | Should -Throw
            }            
        }
    }

    Describe 'Build-ChoiceTable' {
        Context 'when called with a valid choice' {
            It 'should return a table with the choices' {
                [xml]$xml = @"
                <$script:CHOICE_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="choice" $script:ELEMENT_ID="choice" 
                $script:USER_PROMPT="choice prompt" $script:USER_FOOTER_PROMPT="footer prompt" $script:USER_ERROR_PROMPT="error prompt" 
                $script:USER_LOOP_BREAK_SIGNAL="cancel" $script:CHOICE_ALLOW_MULTI_SELECT="false">
                    <$script:CHOICE_ELEMENT $script:CHOICE_ID="1" $script:CHOICE_TEXT="choice1" />
                    <$script:CHOICE_ELEMENT $script:CHOICE_ID="2" $script:CHOICE_TEXT="choice2" />
                </$script:CHOICE_INPUT_ELEMENT>
"@
                $result = Build-ChoiceTable -inputElement $xml.DocumentElement
                $result | Should -BeOfType [hashtable]
                $result.Keys | Should -HaveCount 2
                $result["1"] | Should -Be "choice1"
                $result["2"] | Should -Be "choice2"
            }
        }

        Context 'when called with an invalid choice' {
            It 'should throw an exception if two Choice elements use the same ChoiceID' {
                [xml]$xml = @"
                <$script:CHOICE_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="choice" $script:ELEMENT_ID="choice" 
                $script:USER_PROMPT="choice prompt" $script:USER_FOOTER_PROMPT="footer prompt" $script:USER_ERROR_PROMPT="error prompt" 
                $script:USER_LOOP_BREAK_SIGNAL="cancel" $script:CHOICE_ALLOW_MULTI_SELECT="false">
                    <$script:CHOICE_ELEMENT $script:CHOICE_ID="1" $script:CHOICE_TEXT="choice1" />
                    <$script:CHOICE_ELEMENT $script:CHOICE_ID="1" $script:CHOICE_TEXT="choice2" />
                </$script:CHOICE_INPUT_ELEMENT>
"@
                { Build-ChoiceTable -inputElement $xml.DocumentElement } | Should -Throw
            }

            It 'should throw an exception if one of the Choice elements contains child elements' {
                [xml]$xml = @"
                <$script:CHOICE_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="choice" $script:ELEMENT_ID="choice" 
                $script:USER_PROMPT="choice prompt" $script:USER_FOOTER_PROMPT="footer prompt" $script:USER_ERROR_PROMPT="error prompt" 
                $script:USER_LOOP_BREAK_SIGNAL="cancel" $script:CHOICE_ALLOW_MULTI_SELECT="false">
                    <$script:CHOICE_ELEMENT $script:CHOICE_ID="1" $script:CHOICE_TEXT="choice1">
                        <test />
                    </$script:CHOICE_ELEMENT>
                </$script:CHOICE_INPUT_ELEMENT>
"@
                { Build-ChoiceTable -inputElement $xml.DocumentElement } | Should -Throw
            }

            It 'should throw an exception if one of the elements is not a Choice element' {
                [xml]$xml = @"
                <$script:CHOICE_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="choice" $script:ELEMENT_ID="choice" 
                $script:USER_PROMPT="choice prompt" $script:USER_FOOTER_PROMPT="footer prompt" $script:USER_ERROR_PROMPT="error prompt" 
                $script:USER_LOOP_BREAK_SIGNAL="cancel" $script:CHOICE_ALLOW_MULTI_SELECT="false">
                    <test />
                </$script:CHOICE_INPUT_ELEMENT>
"@
                { Build-ChoiceTable -inputElement $xml.DocumentElement } | Should -Throw
            }
        }
    }

    Describe 'Get-UserChoices' {
        Context 'when called with a valid choice' {
            It 'should return the choice' {
                [xml]$xml = @"
                <$script:CHOICE_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="choice" $script:ELEMENT_ID="choice" 
                $script:USER_PROMPT="choice prompt" $script:USER_FOOTER_PROMPT="footer prompt" $script:USER_ERROR_PROMPT="error prompt" 
                $script:USER_LOOP_BREAK_SIGNAL="cancel" $script:CHOICE_ALLOW_MULTI_SELECT="false">
                    <$script:CHOICE_ELEMENT $script:CHOICE_ID="1" $script:CHOICE_TEXT="choice1" />
                    <$script:CHOICE_ELEMENT $script:CHOICE_ID="2" $script:CHOICE_TEXT="choice2" />
                </$script:CHOICE_INPUT_ELEMENT>
"@               
                Mock Read-Host { return '1' }
                $result = Get-UserChoices -choiceInputChoices @{'1'='choice1';'2'='choice2'} -inputElement $xml.DocumentElement
                $result["choice"] | Should -Be "choice1"
            }

            It 'should return the choice' {
                [xml]$xml = @"
                <$script:CHOICE_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="choice" $script:ELEMENT_ID="choice" 
                $script:USER_PROMPT="choice prompt" $script:USER_FOOTER_PROMPT="footer prompt" $script:USER_ERROR_PROMPT="error prompt" 
                $script:USER_LOOP_BREAK_SIGNAL="cancel" $script:CHOICE_ALLOW_MULTI_SELECT="false">
                    <$script:CHOICE_ELEMENT $script:CHOICE_ID="1" $script:CHOICE_TEXT="choice1" />
                    <$script:CHOICE_ELEMENT $script:CHOICE_ID="2" $script:CHOICE_TEXT="choice2" />
                </$script:CHOICE_INPUT_ELEMENT>
"@                  
                Mock Read-Host { return '2' }
                $result = Get-UserChoices -choiceInputChoices @{'1'='choice1';'2'='choice2'}  -inputElement $xml.DocumentElement
                $result["choice"] | Should -Be "choice2"
            }            

            It 'should throw an exception if user enters break signal' {
                [xml]$xml = @"
                <$script:CHOICE_INPUT_ELEMENT $script:TEMPLATE_DEFINITION_NAME="choice" $script:ELEMENT_ID="choice" 
                $script:USER_PROMPT="choice prompt" $script:USER_FOOTER_PROMPT="footer prompt" $script:USER_ERROR_PROMPT="error prompt" 
                $script:USER_LOOP_BREAK_SIGNAL="cancel" $script:CHOICE_ALLOW_MULTI_SELECT="false">
                    <$script:CHOICE_ELEMENT $script:CHOICE_ID="1" $script:CHOICE_TEXT="choice1" />
                    <$script:CHOICE_ELEMENT $script:CHOICE_ID="2" $script:CHOICE_TEXT="choice2" />
                </$script:CHOICE_INPUT_ELEMENT>
"@                  
                Mock Read-Host { return 'cancel' }
                { Get-UserChoices -choiceInputChoices @{'1'='choice1';'2'='choice2'} -inputElement $xml.DocumentElement } | Should -Throw
            }            
        }
    }
}