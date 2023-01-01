# This also should imply, that this script can only run on Windows, because PowerShell Core version starts with 6.0
#requires -Version 5.1



#region Constants
[string]$script:PLACEHOLDER_PREFIX="!@"
[string]$script:PLACEHOLDER_POSTFIX="@!"
[string]$script:LOOPEND_MARKER="_loopend"
[string]$script:NEW_LINE_IN_WORD="`r"
[string]$script:CONDITION_PREFIX="Condition"
[string]$script:YOUNGER_DATE_CONDITION="ConditionDateYoungerThan"
[string]$script:TEMPLATE_DEFINITION_NAME="Name"
[string]$script:USER_PROMPT="Prompt"
[string]$script:USER_FOOTER_PROMPT="FooterPrompt"
[string]$script:USER_ERROR_PROMPT="ErrorPrompt"
[string]$script:USER_MULTISELECT_PROMPT="MulitselectPrompt"
[string]$script:USER_INPUT_ELEMENT="UserInput"
[string]$script:LOOP_INPUT_ELEMENT="LoopInput"
[string]$script:CHOICE_INPUT_ELEMENT="ChoiceInput"
[string]$script:CHOICE_ELEMENT="Choice"
[string]$script:CHOICE_ID="ChoiceID"
[string]$script:CHOICE_TEXT="ChoiceText"
[string]$script:CHOICE_ALLOW_MULTI_SELECT="AllowMultiselect"
[string]$script:INPUT_ELEMENT_SUFFIX="Input"
[string]$script:PLACEHOLDER_INPUT_ELEMENT="Placeholder"
[string]$script:ELEMENT_ID="ID"
[string]$script:USER_LOOP_BREAK_SIGNAL="BreakKeyword"
[string]$script:VALIDATION_REGEX="ValidateRegex"
[string]$script:TEMPLATE_DEFINITION_FILE_EXTENSION="xml"
[string]$script:TEMPLATE_FILE_EXTENSION="docx"
#endregion Constants

#region Script Variables
[Xml]$Script:TemplateDefinition = $null
[Hashtable]$script:WordTemplateInput = @{}
#endregion Script Variables

#region Private Functions
#region Template Functions
# function that loads every xml file in ./TemplateDefinitions
# adds the value of the "name" attribute of the document element to an array
# returns the array at the end
function Get-MyWordTemplateNames {
    [CmdletBinding()]
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$templatePath = '.\TemplateDefinitions'
    )
    $xmlFiles = Get-ChildItem -Path $templatePath -Filter *.xml
    $names = @()
    foreach ($xmlFile in $xmlFiles) {
        $xml = [xml](Get-Content $xmlFile.FullName)
        $names += $xml.DocumentElement.Name
    }
    return $names
}

#region Get-TemplateDefinitionsPath
function Get-MatchingFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$templateType,
        [Parameter(Mandatory=$true)]
        [string]$templatePath
    )
    $xmlFiles = Get-ChildItem -Path $templatePath -Filter *.xml
    $matchingFiles = @()
    
    foreach ($xmlFile in $xmlFiles) {
        $xml = [xml](Get-Content $xmlFile.FullName)
        if ($xml.DocumentElement.Name -eq $templateType) {
            $matchingFiles += $xmlFile.FullName
        }
    }    

    return $matchingFiles
}

# function that returns the path to the xml file in a folder
# folder is provided as parameter
# name attribute of the xml file document element must match parameter
# should throw an exception if more than one file matches
function Get-TemplateDefinitionsPath {
    [CmdletBinding()]
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$templatePath = '.\TemplateDefinitions',        
        [Parameter(Mandatory=$true)]
        [string]$templateType
    )

    $matchingFiles = @(Get-MatchingFiles -templateType $templateType -templatePath $templatePath)
    if ($matchingFiles.Count -gt 1) {
        throw "More than one file matches the template type $templateType"
    }
    if ($matchingFiles.Count -eq 0) {
        throw "No file matches the template type $templateType"
    }

    # validate if the xml file has a valid schema
    if (-not (Test-MyWordTemplateDefinitionSchema -xmlFilePath $matchingFiles[0])) {
        throw "The xml file $($matchingFiles[0]) does not have a valid schema"
    }

    return $matchingFiles[0]
}
#endregion Get-TemplateDefinitionsPath

function Get-WordTemplatePath {
    [CmdletBinding()]
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$wordTemplatePath,
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]        
        [Parameter(Mandatory=$true)]
        [string]$templateDefinitionFile
    )

    $wordDocuments = Get-ChildItem -Path $wordTemplatePath -Filter *.docx
    $matchingFiles = @()
    foreach ($wordDocument in $wordDocuments) {
        # Get name of the worddocument without extension
        $wordDocumentName = (Split-Path $wordDocument.FullName -Leaf) -replace "\.$($script:TEMPLATE_FILE_EXTENSION)$", ''
        # Get name of the template without extension
        $templateName = (Split-Path $templateDefinitionFile -Leaf) -replace "\.$($script:TEMPLATE_DEFINITION_FILE_EXTENSION)$", ''

        if ($wordDocumentName -eq $templateName) {
            $matchingFiles += $wordDocument.FullName
        }
    }

    if ($matchingFiles.Count -gt 1) {
        throw "More than one file matches the template type $templateType"
    }
    if ($matchingFiles.Count -eq 0) {
        throw "No file matches the template type $templateType"
    }

    Write-Verbose "Found word template $($matchingFiles[0])"
    return $matchingFiles[0]
}
#endregion Template Functions

#region Word Generation Functions
# function that uses the name of the template file and the provided outputpath to create the complete output path for the new word document
# also does name collision detection and adds a counter if collision was detected
# returns the path to the new word document
function Get-DocumentOutputPath {
    [CmdletBinding()]
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$templatePath,
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$outputPath
    )
    $templateFileName = Split-Path $templatePath -Leaf
    $templateFileName = $templateFileName -replace "\.($script:TEMPLATE_DEFINITION_FILE_EXTENSION)$", ''
    $outputFileName = "$($templateFileName).$script:TEMPLATE_FILE_EXTENSION"
    # absolute path is required for the Word api
    $newOutputPath = Join-Path $(Resolve-Path -path $outputPath) $outputFileName
    $counter = 1
    while (Test-Path $newOutputPath) {
        $outputFileName = "$($templateFileName)($($counter)).$script:TEMPLATE_FILE_EXTENSION"
        # absolute path is required for the Word api        
        $newOutputPath = Join-Path $(Resolve-Path -path $outputPath) $outputFileName
        $counter++
    }
    return $newOutputPath
}

function Set-LoopInput {
    param(
        [Parameter(Mandatory=$true)]
        [object]$wordDoc,
        [Parameter(Mandatory=$true)]
        [string]$inputName,
        [Parameter(Mandatory=$true)]
        [Hashtable]$inputTable,
        [string]$itemSeperator=$script:NEW_LINE_IN_WORD
    )

    [string]$inputValue = ""
    # TODO move to separate function
    # ensure the keys are sorted. This works because during input time in the loop every key had a counter prepended for the loop
    $inputTable.GetEnumerator() | Sort-Object -Property key | foreach-object {
        if($_.Name -notlike "*$script:LOOPEND_MARKER") {
            Write-Verbose "Adding value '$($_[0].Value)' for key '$($_[0].Name)' ..."
            $inputValue += "$($_[0].Value) "
        } else {
            $inputValue = $inputValue.Trim()
            $inputValue += $itemSeperator
        }      
    }
    $inputValue = $inputValue.Trim()

    $MatchCase = $false
    $MatchWholeWorld = $true
    $MatchWildcards = $false
    $MatchSoundsLike = $false
    $MatchAllWordForms = $false
    $Forward = $false
    $Wrap = 1
    $Format = $false
    # Replace all occurrences of the text
    $Replace = 2
    Write-Verbose "Replacing placeholder '$script:PLACEHOLDER_PREFIX$($inputName)$script:PLACEHOLDER_POSTFIX' with '$inputValue' in the document body"
    Write-Verbose "Type of input value is $($inputValue.GetType())"          
    $wordDoc.Content.Find.Execute("$script:PLACEHOLDER_PREFIX$($inputName)$script:PLACEHOLDER_POSTFIX", $MatchCase, $MatchWholeWorld, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $inputValue, $Replace) | Out-Null
    Write-Verbose "Loop input is not suported for header and footer"
}

function Set-TextInput {
    param(
        [Parameter(Mandatory=$true)]
        [object]$wordDoc,
        [Parameter(Mandatory=$true)]
        [string]$inputName,
        [Parameter(Mandatory=$true)]
        [string]$inputValue
    )

    $MatchCase = $false
    $MatchWholeWorld = $true
    $MatchWildcards = $false
    $MatchSoundsLike = $false
    $MatchAllWordForms = $false
    $Forward = $false
    $Wrap = 1
    $Format = $false
    # Replace all occurrences of the text
    $Replace = 2  
    Write-Verbose "Replacing placeholder $script:PLACEHOLDER_PREFIX$($inputName)$script:PLACEHOLDER_POSTFIX with $inputValue in the document body"          
    $wordDoc.Content.Find.Execute("$script:PLACEHOLDER_PREFIX$($inputName)$script:PLACEHOLDER_POSTFIX", $MatchCase, $MatchWholeWorld, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $inputValue, $Replace) | Out-Null
    Write-Verbose "Replaceing $script:PLACEHOLDER_PREFIX$($inputName)$script:PLACEHOLDER_POSTFIX with $inputValue in the document header"
    $header = $wordDoc.Sections.Item(1).Headers.Item(1)
    $header.Range.Find.Execute("$script:PLACEHOLDER_PREFIX$($inputName)$script:PLACEHOLDER_POSTFIX", $MatchCase, $false, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $true, $Wrap, $Format, $inputValue, $Replace) | Out-Null  
}

# function that creates a word document from a template
# returns the path to the created word document
function New-MyWordDocument {
    [CmdletBinding()]
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$templatePath,
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$wordTemplatePath,
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$outputPath,
        [hashtable]$templateInput
    )
    $word = $null
    [string]$newFilePath = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $wordTemplatePath = Resolve-Path -Path $wordTemplatePath
        Write-Verbose "Opening template $wordTemplatePath"
        $wordDoc = $word.Documents.Open($wordTemplatePath)
        $wordDoc.Activate() | Out-Null
        foreach ($key in $templateInput.Keys) {
            if($templateInput[$key] -is [string]) {
                Set-TextInput -wordDoc $wordDoc -inputName $key -inputValue $templateInput[$key]
            # this will handle inputs from Loop elements, and from ChoiceInput elements which allow multiple values
            } elseif ($templateInput[$key] -is [hashtable] -or $templateInput[$key] -is [ordered]) {
                Set-LoopInput -wordDoc $wordDoc -inputName $key -inputTable $templateInput[$key]
            } else {
                throw "Invalid input type $($templateInput[$key].GetType())"
            }
        }
        $newFilePath = Get-DocumentOutputPath -templatePath $templatePath -outputPath $outputPath
        Write-Verbose "Saving document to $newFilePath"    
        $wordDoc.SaveAs($newFilePath) | Out-Null
    } catch {   
        $newFilePath = $null 
        throw $_
    } finally {
        if ($null -ne $word) {
            $wordDoc.Close()       
            $word.Quit()
            # Ensure that COM objects are released
            # if you don't do this, you may get an error from Word when you try to normally open the template word document 
            # this will also prevent you from using that template word docuemnt in this script again in headless Word mode
            # because the invisble Word will inform you that an error occured the last time you tried to open the template
            # and, because Word is invislbe, you won't be able to acknowledge the error which then in turn sets the Word
            # instance in the backround into an idle state.
            $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
            [gc]::Collect()
            [gc]::WaitForPendingFinalizers()            
        }
    }

    return $newFilePath
}
#endregion Word Generation Functions

#region Validation Functions
# function that checks the xml element attributes of a provided attribute for any attribute starting with 'Condition'
# if an attribute starting with 'Condition' is found, it returns true
# if no attribute starting with 'Condition' is found, it returns false
function Test-HasCondition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$xmlElement
    )
    $hasCondition = $false
    foreach ($attribute in $xmlElement.Attributes) {
        if ($attribute.Name -like "$script:CONDITION_PREFIX*") {
            $hasCondition = $true
            break
        }
    }
    return $hasCondition
}

function Test-UserInputCondition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$inputString,
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$inputElement
    )    

    $conditionMet = $true

    $dateTableKey = $inputElement.Attributes[$script:YOUNGER_DATE_CONDITION].'#text'
    if($dateTableKey) {
        $dateString = $inputString
        # Use $script:WordTemplateInput as hashtable to store the already provided input
        $dateYoungerThan = Test-InputDate -dateString $dateString -dateTable $script:WordTemplateInput -dateTableKey $dateTableKey
        if (-not $dateYoungerThan) {
            Write-Host "The date $dateString is not younger than the date $($script:WordTemplateInput[$dateTableKey])" -ForegroundColor Yellow
            $conditionMet = $false
        }
    }

    return $conditionMet
}

# function that validates a string to be a valid date and that the date is younger than a date provided in a hashtable as parameter
# returns true if the string is a valid date and younger than the date provided in the hashtable
# returns false if the string is not a valid date or if the date is not younger than the date provided in the hashtable
# throws an exception if one of the provided dates is not a valid date
function Test-InputDate {
    [CmdletBinding()]
    param(
        [ValidateScript({$null -ne [DateTime]::Parse($_, [System.Globalization.CultureInfo]::CurrentCulture)})]
        [Parameter(Mandatory=$true)]
        [string]$dateString,
        [Parameter(Mandatory=$true)]
        [hashtable]$dateTable,
        [Parameter(Mandatory=$true)]
        [string]$dateTableKey
    )
    $date = $null
    Write-Verbose "Validating younger date '$dateString'"
    $date = [DateTime]::Parse($dateString, [System.Globalization.CultureInfo]::CurrentCulture)
    if ($null -eq $date) {
        return $false
    }
    Write-Verbose "Validating older date '$($dateTable[$dateTableKey])' for the id '$dateTableKey'"
    $olderDate = [DateTime]::Parse($dateTable[$dateTableKey], [System.Globalization.CultureInfo]::CurrentCulture)
    if ($date -lt $olderDate) {
        return $false
    }
    return $true
}

#region Schema Validation Functions
# function that validates xml against schema
# returns true if xml is valid
# returns false if xml is invalid
function Test-MyWordTemplateDefinitionSchema {
    [CmdletBinding()]
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$xsdPath='.\TemplateDefinitions\TemplateDefinition.xsd',
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [Parameter(Mandatory=$true)]
        [string]$xmlFilePath
    )
    $xmlReaderSettings = New-Object System.Xml.XmlReaderSettings
    $xmlReaderSettings.Schemas.Add('', $xsdPath) | Out-Null
    $xmlReaderSettings.ValidationType = [System.Xml.ValidationType]::Schema
    try {
        $xmlReader = [System.Xml.XmlReader]::Create($xmlFilePath, $xmlReaderSettings)
        try {
            while ($xmlReader.Read()) {}
        }
        catch {
            Write-Host "Template defintion file '$xmlFilePath' invalid." -ForegroundColor Red
            Write-Host "`t$($_.Exception.Message)"
            return $false
        }
    } finally {
        if ($null -ne $xmlReader) {
            $xmlReader.Close()
        }
    }

    return $true
}

# function that validates that the MyWordTemplateDefinition element has a Name attribute where the value matches the filename of the template definition file
# returns true if the Name attribute matches the filename
# returns false if the Name attribute does not match the filename
function Test-TemplateDefinitionFilename {
    [CmdletBinding()]
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [Parameter(Mandatory=$true)]
        [string]$templateDefinitionFilePath
    )
    [bool]$valid = $true

    # validate the Name attribute of the TemplateDefinition element against the filename
    $xml = [xml](Get-Content -path $templateDefinitionFilePath)
    $templateDefinitionName = $xml.MyWordTemplateDefinition.Attributes[$script:TEMPLATE_DEFINITION_NAME].'#text'
    if ($templateDefinitionName -ne (Split-Path -path $templateDefinitionFilePath -LeafBase)) {
        Write-Host "Template defintion file '$($templateDefinitionFilePath)' invalid." -ForegroundColor Red
        Write-Host "`tThe Name attribute of the MyWordTemplateDefinition element does not match the filename."
        $valid = $false
    }   

    return $valid
}

# function which validates every xml file in template defininitions folder against schema
# returns true if all xml files are valid
# returns false if at least one xml file is invalid
function Test-MyWordTemplateDefinitions {
    [CmdletBinding()]
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$templateDefinitionPath = '.\TemplateDefinitions'
    )
    $xmlFiles = Get-ChildItem -Path $templateDefinitionPath -Filter *.xml
    $invalidSchemaFiles = @()
    foreach ($xmlFile in $xmlFiles) {
        if (-not (Test-MyWordTemplateDefinitionSchema -xmlFilePath $xmlFile.FullName)) {
            $invalidSchemaFiles += $xmlFile.FullName
        }

        # validate the Name attribute of the TemplateDefinition element against the filename
        if (-not (Test-TemplateDefinitionFilename -templateDefinitionFilePath $xmlFile)) {
            $invalidSchemaFiles += $xmlFile.FullName
        }
    }
    return $invalidSchemaFiles
}
#endregion Schema Validation Functions

function Test-WordInstallation {
    # Validate that Word is installed for the current user
    $word = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "Microsoft 365*" }
    if (-not $word) {
        return $false
    } else {
        return $true
    }
}

function Get-FileNamesInFolder {
    [CmdletBinding()]
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$folderPath,
        [string]$filter = '*'
    )
    $files = Get-ChildItem -Path $folderPath -Filter $filter | Select-Object Name
    return $files
}

function Test-WordTemplatesAgainstWordTemplateDefintions {
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$templateDefinitionsPath = '.\TemplateDefinitions',
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$templatesPath = '.\Templates'
    )
   # ensure that there is a word document in the Templates folder for every template defintion in the template definitions folder where the name of the word document is the same as the name of the template definition
   $templateDefinitions = Get-FileNamesInFolder -folderPath $templateDefinitionsPath -Filter *.xml
   $templates = Get-FileNamesInFolder -folderPath $templatesPath -Filter *.docx
   foreach ($templateDefinition in $templateDefinitions) {
       $templateDefinitionName = $templateDefinition.Name -replace "\.$($script:TEMPLATE_DEFINITION_FILE_EXTENSION)$"
       $template = $templates | Where-Object { $($_.Name -replace "\.$($script:TEMPLATE_FILE_EXTENSION)$") -eq $templateDefinitionName }
       if (-not $template) {
           Write-Host "Word template '$templatesPath\$templateDefinition.$script:TEMPLATE_FILE_EXTENSION' is missing." -ForegroundColor Red
           return $false
       }
   }
   return $true
}

# function that checks if a given word template contains placeholders, 
# surrounded by $script:PLACEHOLDER_PREFIX and $script:PLACEHOLDER_POSTFIX, 
# to all element ids in the matching template definition
# returns true if all placeholders are present
# returns false if at least one placeholder is missing
function Test-WordTemplatePlaceholders {
    [CmdletBinding()]
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [Parameter(Mandatory=$true)]
        [string]$templatePath,
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$templateDefinitionsPath = '.\TemplateDefinitions'
    )
    Write-Warning "Test-WordTemplatePlaceholders not implemented yet."
    # TODO: implement
}
#endregion Validation Functions

#region Get-TemplateInput
#region Get-ChoiceInput
function Get-UserChoicesAsTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$choiceInputElementId,
        # Hashtable with all the allowed choices
        [Parameter(Mandatory=$true)]
        [hashtable]$allowedChoices,
        # This can be a single int in a string
        # or a comma separated list of ints in a string if $isMultiSelect is true
        # REMARKS: We assume here, that the user input already has been validated before calling this function
        [Parameter(Mandatory=$true)]
        [string]$stringWithUserChoices,
        [Parameter(Mandatory=$true)]
        [bool]$isMultiSelect
    )
    $choices = @{}

    if(-not $isMultiSelect) {
        # this will throw an exception if the function has been called with a comma seperated user input
        $allowedChoiceID = $allowedChoices.Keys | Where-Object { $_ -eq [int]$stringWithUserChoices }          
        $choices.Add($choiceInputElementId, $allowedChoices[$allowedChoiceID])
    } else {
        [int]$counter = 1
        $stringWithUserChoices -split "," | ForEach-Object {    
            $selectedValue = $allowedChoices[$_]
            if($choices.ContainsValue($selectedValue)) {
                throw "Multiple selections of the same value are not allowed."
            }
            $choices.Add("$choiceInputElementId$counter", $selectedValue)
            $counter++
        }
    }   
    
    return $choices
}

function Test-UserChoice {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [bool]$isMultiSelect,
        [Parameter(Mandatory=$true)]
        [string]$choiceInputElementChoice,
        [Parameter(Mandatory=$true)]
        [array]$allowedChoiceIDs
    )
    $valid = $true

    if (-not $isMultiSelect) {
        $valid = ($choiceInputElementChoice -match "^[0-9]+$") -and ($choiceInputElementChoice -in $allowedChoiceIDs)
    } elseif ($isMultiSelect -and $choiceInputElementChoice -match "^[0-9]+(,[0-9]+)*$") {
        # check if all choices are unique
        $choiceInputElementChoices = $choiceInputElementChoice -split ","
        $uniqueChoices = $choiceInputElementChoices | Select-Object -Unique
        if ($uniqueChoices.Count -eq $choiceInputElementChoices.Count) {
            $uniqueChoices | ForEach-Object {
                if (-not ($_ -in $allowedChoiceIDs)) {
                    $valid = $false
                    break
                }
            }
        } else {
            $valid = $false
        }
    }

    return $valid
}


function Get-UserChoices {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$inputElement,
        [Parameter(Mandatory=$true)]
        [hashtable]$choiceInputChoices
    ) 
    $choices = @{}

    # prompt the user to select one of the choices
    $choiceInputElementId = $inputElement.Attributes[$script:ELEMENT_ID].'#text' 

    $choiceInputElementPrompt = $inputElement.Attributes[$script:USER_PROMPT].'#text' + "`r`n"
    # Sort the Names of the hashtable keys. 
    # Normally we build the hashtable in ordered fashion, but the ordering is lost when handed over to this function which can only happen as hashtable
    $choiceInputElementPrompt += $choiceInputChoices.Keys | Sort-Object -Property Name | ForEach-Object { "`r$($choiceInputChoices[$_]) ($_)`n" }
    $choiceBreakSingnal = $inputElement.Attributes[$script:USER_LOOP_BREAK_SIGNAL].'#text'
    Write-Host $choiceInputElementPrompt

    [bool]$isMultiSelect = [System.Convert]::ToBoolean($inputElement.Attributes[$script:CHOICE_ALLOW_MULTI_SELECT].'#text')
    if($isMultiSelect) {
        Write-Host $inputElement.Attributes[$script:USER_MULTISELECT_PROMPT].'#text' + "`r`n"
    }

    # check if the user entered a valid choice
    [bool]$choiceValid = $false
    do {
        $choiceInputElementChoice = "$($inputElement.Attributes[$script:USER_FOOTER_PROMPT].'#text') ('$choiceBreakSingnal' to stop script)"
        $choiceInputElementChoice = Read-Host $choiceInputElementChoice 
        # exit script execution if the user entered break keyword
        if ($choiceInputElementChoice -eq $choiceBreakSingnal) {
            throw "User canceled script execution."
        }

        $choiceValid = Test-UserChoice -isMultiSelect $isMultiSelect -choiceInputElementChoice $choiceInputElementChoice -allowedChoiceIDs $choiceInputChoices.Keys
        $choiceInputElementChoice = Get-UserChoicesAsTable -choiceInputElementId $choiceInputElementId -allowedChoices $choiceInputChoices -stringWithUserChoices $choiceInputElementChoice -isMultiSelect $isMultiSelect
    } while (-not $choiceValid -or $null -eq $choiceInputElementChoice -or $choiceInputElementChoice.Count -eq 0)
    $choices = $choiceInputElementChoice

    return $choices
}

function Build-ChoiceTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$inputElement
    )    
    $choiceInputChoices = [ordered]@{}

    $choiceInputElementId = $inputElement.Attributes[$script:ELEMENT_ID].'#text'
    foreach ($choiceElement in $inputElement.ChildNodes) {
        if($choiceElement -is [System.Xml.XmlElement]) {
            if($choiceElement.LocalName -ne $script:CHOICE_ELEMENT) {
                throw "Child element '$($choiceElement.LocalName)' of choice input element '$choiceInputElementId' is not allowed."
            }
            # throw an error if the child contains other elements
            if($choiceElement.ChildNodes.Count -gt 0) {
                throw "Child element '$($choiceElement.LocalName)' of choice input element '$choiceInputElementId' contains other elements."
            }
            $choiceElementId = $choiceElement.Attributes[$script:CHOICE_ID].'#text'
            $choiceElementText = $choiceElement.Attributes[$script:CHOICE_TEXT].'#text'
            Write-Verbose "Adding choice '$choiceElementText' with id '$choiceElementId' to choice input element '$choiceInputElementId'"
            $choiceInputChoices.Add($choiceElementId, $choiceElementText)
        }
    }
    
    return $choiceInputChoices
}

# function that takes a xml element with the local name "ChoiceInput" with any number of direct child elements of type "Choice"
# the "ChoiceInput" element has an attribute "ID" which is the id of the element in the word template
# the "ChoiceInput" elements have an attribute "Prompt" which is used to prompt the user to select one of the choices
# the "Choice" elements have an attribute "ChoiceID" which is the id of the choice which is a number
# the "Choice" elements have an attribute "ChoiceText" which is the text of the choice which is a string
# It prompts the user to select one of the choices by displaying the "Prompt" attribute of the "ChoiceInput" element and listing all "ChoiceText" attributes of the "Choice" elements
# behind each "ChoiceText" attribute it displays the "ChoiceID" attribute of the "Choice" element in brackets
# It returns a hashtable with the key being the id of the element in the word template and the value being the selected "ChoiceText"
# only one choice can be selected
function Get-ChoiceInput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$inputElement
    )
    $templateInput = @{}
  
    $choiceInputChoices = Build-ChoiceTable -inputElement $inputElement

    $templateInput = Get-UserChoices -inputElement $inputElement -choiceInputChoices $choiceInputChoices

    return $templateInput
}
#endregion Get-ChoiceInput

#region Get-LoopInput
function Get-LoopChild {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$loopElement,
        [Parameter(Mandatory=$true)]
        [string]$loopElementId,
        [Parameter(Mandatory=$true)]
        [string]$breakkeyword,
        [Parameter(Mandatory=$true)]
        [int]$loopCounter
    )
    $templateInput = [ordered]@{}

    foreach ($loopChildElement in $loopElement.ChildNodes) {
        if($loopChildElement -is [System.Xml.XmlElement]) {
            if($loopChildElement.LocalName -ne $script:USER_INPUT_ELEMENT) {
                throw "Child element '$($loopChildElement.LocalName)' of loop element '$loopElementId' is not supported."
            }
            $loopChildElement.Attributes[$script:USER_PROMPT].'#text' = "`t`t$($($loopChildElement.Attributes[$script:USER_PROMPT].'#text').Trim()))"                
            Write-Verbose "Prompt attribute of child element '$($loopChildElement.Name)' is '$($loopChildElement.Attributes[$script:USER_PROMPT].'#text')'"
            # get the user input for the current child element                
            $userInput = Invoke-TemplateConfigElement -templateElement $loopChildElement
            if($userInput.Keys.Count -gt 1) {
                throw "More than one user input returned for element '$($loopChildElement.Name)'. This is currently not supported."
            }
            if($userInput.Keys.Count -eq 0) {
                throw "No user input returned for element '$($loopChildElement.Name)'."
            }
            # prepend the loopCounter value to the element id
            [string]$elementId = "$loopCounter$($userInput.Keys[0])"
            $userEntry = $userInput.Values[0]
            Write-Verbose "User input for element '$($loopChildElement.Name)' is '$userEntry'"
            if(-not ($userEntry -eq $breakkeyword)) {
                # add the user input to the iteration input hashtable
                $templateInput.Add($elementId, $userEntry)
            } else {
                Write-Verbose "Input is '$userEntry'. Breaking input loop."
                $templateInput.Add($breakkeyword, $breakkeyword)
                break
            }
        }
    }

    return $templateInput
}

# function that takes an xml element with the local name "loop" and iterates over its child elements
# for every child element it calls Invoke-TemplateConfigElement which returns a hashtable with the element id as key and the user input as value
# it appends a counter value to the end of the element id for every iteration
# when it is done it returns a hashtable with the element id as key and the hashtable with the user input as value
function Get-LoopInput{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$inputElement
    )
    $templateInput = [ordered]@{}

    # get the element id of the loop element
    $loopElementId = $inputElement.Attributes[$script:ELEMENT_ID].'#text'
    Write-Verbose "Loop element id: '$loopElementId'"

    # get the number of iterations
    $counter = 1
    $breakkeyword = $inputElement.Attributes[$script:USER_LOOP_BREAK_SIGNAL].'#text'
    # write Prompt attribute of the loop element to screen with the text "Start - " prepended and "($breakkeword to stop)" attribute appended
    Write-Host "Start - $($inputElement.Attributes[$script:USER_INPUT_ELEMENT].'#text') ('$breakkeyword' to stop)"
    # iterate over the child elements
    [bool]$breakLoop = $false
    do {    
        Write-Verbose "Iterating over loop. Iteration $counter"
        $templateInput += Get-LoopChild -loopElement $inputElement -loopElementId $loopElementId -breakkeyword $breakkeyword -loopCounter $counter
        $breakLoop = $templateInput.Keys -contains $breakkeyword
        if($breakLoop) {
            $templateInput.Remove($breakkeyword)
        } else {
            # Prepend an outer loop end marker so that later we can discern between complete loop iterations
            $templateInput.Add("$($counter)$script:LOOPEND_MARKER", $script:NEW_LINE_IN_WORD)
        }
        $counter += 1        
    } while (-not $breakLoop)    
    # write "End - " to screen and the Prompt attribute of the loop element
    Write-Host "End - $($inputElement.Attributes[$script:USER_PROMPT].'#text')"
        
    return $templateInput
}
#endregion Get-LoopInput

function Get-UserInput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$inputElement
    )
    $userInput = @{}
    Write-Verbose "Prompting user for input (Prompt: '$($inputElement.Prompt)') ..."
    $userProvidedInput = ""
    do { # The outer loop ensures that the input condition is met
        if ($inputElement.Attributes[$script:VALIDATION_REGEX]) {
            do { # The inner loop ensures that the input matches the regex
                Write-Verbose "Input with regex validation (Regex: $($inputElement.Attributes[$script:VALIDATION_REGEX].'#text')) ..."
                $userProvidedInput = Read-Host "$($inputElement.Prompt)"
                Write-Verbose "Input was: '$userProvidedInput'"                
            } while (-not ($userProvidedInput -match $inputElement.Attributes[$script:VALIDATION_REGEX].'#text'))
        }
        else {
            $userProvidedInput = Read-Host $inputElement.Prompt
        }
    } while (-not (Test-UserInputCondition -inputElement $inputElement -inputString $userProvidedInput))
    $userInput += @{ $inputElement.ID = $userProvidedInput }
    return $userInput
}

# function that takes a hashtable of user inputs
# it checks every entry in the hash table agaings $script:WordTemplateInput
# if the entry is not in $script:WordTemplateInput, it is added to the hashtable
# if the entry is in $script:WordTemplateInput, it is checked if the value is the same
# if the value is the same, nothing is done
# if the value is different, the user is asked if the value should be overwritten
# if the user answers yes, the value is overwritten
# if the user answers no, the value is not overwritten
# if the user answers cancel, the function returns $null
function Add-TemplateInput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$userInput
    )
    Write-Verbose "Making a local copy of script wide template input table."
    $templateInput = $script:WordTemplateInput
    
    foreach ($key in $userInput.Keys) {
        if ($script:WordTemplateInput.ContainsKey($key)) {
            if ($script:WordTemplateInput[$key] -ne $userInput[$key]) {
                $overwrite = Read-Host "The value for '$key' is already set to '$($script:WordTemplateInput[$key])'. Do you want to overwrite it with '$($userInput[$key])' (y for yes, n for no, c for cancel)?"
                if ($overwrite -eq 'y') {
                    Write-Verbose "Overwriting value for '$key' with '$($userInput[$key])'."
                    $templateInput[$key] = $userInput[$key]
                } elseif ($overwrite -eq 'n') {
                    # do nothing
                } elseif ($overwrite -eq 'c') {
                    return $false
                } else {
                    throw "Invalid input '$overwrite'."
                }
            }
        } else {
            $templateInput.Add($key, $userInput[$key])
        }
    }

    Write-Verbose "Updating script wide template input table."
    $script:WordTemplateInput = $templateInput
    # return true if the function was successful
    return $true
}

function Invoke-TemplateConfigElement
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$templateElement
    )
    $templateInput = @{}
    
    if($null -ne $templateElement.ID -and $script:WordTemplateInput.ContainsKey($templateElement.ID)) {
        Write-Verbose "Element '$($templateElement.ID)' already exists in the template input hashtable. Skipping."
        return $templateInput
    }

    # if the local name of the element ends with $script:INPUT_ELEMENT_SUFFIX
    # call a function with the name "Get-<LocalName>" and the parameter inputElement which takes $templateElement
    # the function should return a hashtable
    # the hashtable should contain a single key value pair
    # the key should be the ID of the element
    # the value can be anything
    # the hashtable is added to the $templateInput hashtable
    # if the local name of the element does not end with $script:INPUT_ELEMENT_SUFFIX but is $script:PLACEHOLDER_INPUT_ELEMENT
    # the element is added to the $templateInput hashtable
    # the key is the ID of the element
    # the value is the inner text of the element
    if ($templateElement.LocalName -like "*$script:INPUT_ELEMENT_SUFFIX") {
        $functionName = "Get-" + $templateElement.LocalName
        $result = & $functionName -inputElement $templateElement
        if($result -is [Hashtable]) {
            $templateInput += $result
        } else {
            $templateInput += @{ $templateElement.ID = $result }
        }
    } elseif ($templateElement.LocalName -eq $script:PLACEHOLDER_INPUT_ELEMENT) {
        $templateInput += @{ $templateElement.ID = $templateElement.InnerText }
    } else {
        Write-Verbose "Element '$($templateElement.LocalName)' is not an input element. Skipping."
    }

    return $templateInput
}

function Get-TemplateInputRecursive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$xml
    )
    $templateInput = @{}
    $templateInput += Invoke-TemplateConfigElement -templateElement $xml

    # REMARKS: Loop childs do not support other nested loops and they handle their own child elements, so we can skip their childs.
    if ($xml.LocalName -eq $script:LOOP_INPUT_ELEMENT) {
        return $templateInput
    }

    $xml.ChildNodes | ForEach-Object {

        if ($_ -is [System.Xml.XmlElement]) {
            # call Get-TemplateInput with child node
            # add returned hashtable to templateInput
            $result = Get-TemplateInputRecursive -xml $_
            # if the result is not already in $templateInput
            # add it
            $result.Keys | ForEach-Object {
                if (-not $templateInput.ContainsKey($_)) {
                    Write-Verbose "Adding $_ to templateInput ..."
                    $templateInput.Add($_, $result[$_])
                }
                if (-not $script:WordTemplateInput.ContainsKey($_) ) {
                    Write-Verbose "Adding $_ to script wide input hashtable ..."
                    $script:WordTemplateInput.Add($_, $result[$_])
                }
            }
        }
    }

    return $templateInput
}

function Get-TemplateInput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$xml
    )
    return Get-TemplateInputRecursive -xml $xml
}
#enregion Get-TemplateInput
#endregion Private Functions

#region Public Functions
function Get-MyWordTemplate {
    [CmdletBinding()]
    param( 
        # REMARKS: Validation of the templyte type is disabled because it is not possible to pass the templatePath
        # parameter that allows only values which are returned by Get-MyWordTemplateNames    
        # [ValidateScript({$_ -in $(Get-MyWordTemplateNames -templatePath $templatePath)})]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]            
        [array]$templateType,
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$templatePath = '.\TemplateDefinitions',
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$wordTemplatePath = '.\Templates',
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$outputpath = '.\GeneratedDocuments'
    )
    <#
    .SYNOPSIS
        Generates a word document based on a word document template in the Templates folder and a template definition file in the TemplateDefinitions folder.
    .DESCRIPTION
        Generates a word document based on a word document template in the Templates folder and a template definition file in the TemplateDefinitions folder.
        The word document template contains placeholders which are replaced with the values specified in the template definition file.
        Also, the template definition file can contain user input prompts which are used to prompt the user for input.
        So the caller can interactively specify the values for the defined placeholders.
    .PARAMETER templateType
        The name of the template definition file in the TemplateDefinitions folder.
        The name of the word document template in the Templates folder must be the same as the name of the template definition file.
    .PARAMETER templatePath
        The path to the TemplateDefinitions folder.
        The default value is '.\TemplateDefinitions'.
        The definition must adhere to the schema 'TemplateDefinition.xsd' in the TemplatesDefinitions folder.
    .PARAMETER wordTemplatePath
        The path to the word templates folder.
        The default value is '.\Templates'.
    .EXAMPLE
        Get-MyWordTemplate -templateType 'MyTemplate'
        Generates a word document based on the template definition file 'MyTemplate.xml' in the TemplateDefinitions folder and the word document template 'MyTemplate.docx' in the Templates folder.
    .EXAMPLE
        Get-MyWordTemplate -templateType 'MyTemplate' -templatePath 'C:\MyTemplateDefinitions'
        Generates a word document based on the template definition file 'MyTemplate.xml' in the folder 'C:\MyTemplateDefinitions' and the word document template 'MyTemplate.docx' in the Templates folder.
    .EXAMPLE
        Get-MyWordTemplate -templateType 'MyTemplate' -templatePath 'C:\MyTemplateDefinitions' -wordTemplatePath 'C:\MyWordTemplates'
        Generates a word document based on the template definition file 'MyTemplate.xml' in the folder 'C:\MyTemplateDefinitions' and the word document template 'MyTemplate.docx' in the folder 'C:\MyWordTemplates'.
    .NOTES
        Requires the Microsoft Word Object Library to be installed.
    #>
    begin {$script:WordTemplateInput = @{}}
    process {   
        $templateType | ForEach-Object {
            # first get the path to the template definition file
            $templateDefintionPath = Get-TemplateDefinitionsPath -templateType $_ -templatePath $templatePath
            Write-Verbose "Path to template definition file is '$templateDefintionPath' ..."
            # then load the template definition file
            $script:TemplateDefinition = [xml](Get-Content $templateDefintionPath)
            # then get the word template input
            # $script:WordTemplateInput will be filled  by Add-TemplateInput
            # Add-TemplateInput returns $true if all input could be added or updated
            if(Add-TemplateInput -userInput $(Get-TemplateInput -xml $script:TemplateDefinition.DocumentElement)) {
                # then load the word document     
                $wordTemplateFilePath = Get-WordTemplatePath -templateDefinitionFile $templateDefintionPath -wordTemplatePath $wordTemplatePath           
                New-MyWordDocument -templatePath $templateDefintionPath -wordTemplatePath $wordTemplateFilePath -templateInput $script:WordTemplateInput -outputPath $outputpath | Out-Null
            }
        }
    }
}

function Test-MyWordTemplateInstallation {
    [CmdletBinding()]
    param(
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$templateDefinitonsPath = '.\TemplateDefinitions',
        [string]$templatesPath = '.\Templates'
    )
    <#
    .SYNOPSIS
    Tests if prerequisites for the Get-MyWordTemplate module are met.
    .DESCRIPTION
    Tests if word is installed, if the template definitions are valid and if the word templates are installed.
    .PARAMETER templateDefinitonsPath
    The path to the folder containing the template definitions.
    .PARAMETER templatesPath
    The path to the folder containing the word templates.
    .EXAMPLE
    Test-MyWordTemplateInstallation
    .EXAMPLE
    Test-MyWordTemplateInstallation -templateDefinitonsPath '.\TemplateDefinitions' -templatesPath '.\Templates'
    #>
    # first check if Word is installed
    if (-not (Test-WordInstallation)) {
        throw "Word is not installed."
    }

    $invalidTemplates = Test-MyWordTemplateDefinitions -templateDefinitionPath $templateDefinitonsPath
    # then check if the template definitions are valid
    if ($invalidTemplates.Count -gt 0) {
        $invalidTemplates | ForEach-Object {
            Write-Warning "Template definition file '$($_)' is not valid."
        }
        throw "One or more word template definitions are not valid."
    }
    
    # then check if the word templates are installed
    if (-not (Test-WordTemplatesAgainstWordTemplateDefintions -templateDefinitionsPath $templateDefinitonsPath -templatesPath $templatesPath)) {
        throw "One or more word templates are missing."
    }

    # should be ok because we check with "requires 5.1" at the beginning of the script
    if($Env:OS -ne "Windows_NT") {
        throw "This script can only run on Windows"
    }    

    Write-Verbose "Prequisites for Get-MyWordTemplate are met."
    return $true
}

# function that validates a word template file
function Test-WordTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$templatePath,
        # validate that path is a valid path
        [ValidateScript({Test-Path $_})]
        [string]$templateDefinitionsPath = '.\TemplateDefinitions'
    )
    <#
    .SYNOPSIS
    Tests if a word template file is valid.
    .DESCRIPTION
    Tests if a word template file is valid.
    Tests if all placeholders in the word template file are defined in the matching template definition file.
    .PARAMETER templatePath
    The path to the word template file.
    .PARAMETER templateDefinitionsPath
    The path to the folder containing the template definitions.
    This parameter is optional.
    .EXAMPLE
    Test-WordTemplate -templatePath 'C:\MyTemplate.docx'
    .NOTES
    Assumes that the template definition file has the same name as the word template file.
    #>
    # first check if the file exists
    if (-not (Test-Path $templatePath)) {
        throw "File '$templatePath' does not exist."
    }

    # then check if word template contains placeholders to all element ids in the matching template definition
    if (-not (Test-WordTemplatePlaceholders -templatePath $templatePath -templateDefinitionsPath $templateDefinitionsPath)) {
        throw "File '$templatePath' is not a word template."
    }

    return $true
}
#endregion Public Functions

#region Exports
Export-ModuleMember -Function Get-MyWordTemplate
Export-ModuleMember -Function Test-MyWordTemplateInstallation
#endregion Exports