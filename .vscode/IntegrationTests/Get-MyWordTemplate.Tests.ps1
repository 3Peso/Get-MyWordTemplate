Import-Module .\Get-MyWordTemplate.psm1 -Force

InModuleScope Get-MyWordTemplate {
    BeforeAll {
        #$script:verbPref = $VerbosePreference
        #$VerbosePreference = 'Continue'            
    }

    AfterAll {
        #$VerbosePreference = $script:verbPref
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
            
            It 'should create a word document' {     
                # mock Test-Path
                Mock Test-Path { return $false } -ParameterFilter { $path -eq '.\Tests\GeneratedDocuments\protocol.docx' }
                Mock Test-Path { return $true }
                Mock Get-DocumentOutputPath { return Join-Path $(Resolve-Path -path '.\Tests\GeneratedDocuments') 'protocol.docx' }
                Mock Test-MyWordTemplateDefinitionSchema { return $true }
                $choiceInput = [ordered]@{choice1 = 'entry 1'
                                          choice2 = 'entry 2'}
                $templateInput = @{test = 'TEST'
                                   choice = $choiceInput}
                $result = New-MyWordDocument -templatePath ".\Tests\validtemplatedefinitions\protocol.xml" -wordTemplatePath ".\Tests\validtemplates\protocol.docx" -outputPath ".\Tests\GeneratedDocuments" -templateInput $templateInput
                $result | Should -Be $true
            }            
            
            AfterEach {
                $testfile = Resolve-Path -path '.\Tests\GeneratedDocuments\*.docx' -ErrorAction SilentlyContinue
                if (Test-Path -Path $testfile) {
                    Remove-Item -Path $testfile -Force -ErrorAction SilentlyContinue
                } else {
                    Write-Host "File '$testfile' does not exist." -ForegroundColor Yellow
                }
            }
        }
    }

    Describe 'Get-MyWordTemplate' {
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
}
