# Get-MyWordTemplate
## Purpose
### Personal Intro
I developed this script to autmate some work related processes of mine, creating word documents time and time again, following the same patterns all the time, only with small overall differences. Sure, I could have gone the way of Word macros. But the can pose a security risk, and it never came to my mind to learn how to make them. There are better things to learn in life ;-)
### General Intro
`Get-MyWordTemplate` basically takes an XML file, interprets the elements of the file by asking the user to provide input. It then takes this input and replaces placeholders in the corresponding Word template file. This template file can already be in the corret format, with headers, fonts, tables, etc. Only the content will be inserted.
## How does it work?
First you have to provide an XML template definition in the local `.\TemplateDefinitions`-folder. An example for such a XML defintion could be the following:
```
<MyWordTemplateDefinition Name="protocol">
    <UserProvided>
        <UserInput Name="Test" ID="test" Prompt="Enter Date" />
        <ChoiceInput Name="Choice" Prompt="Choice" ID="choice" 
            BreakKeyword="cancel" FooterPrompt="Choose" ErrorPrompt="Not allowed"
            AllowMultiselect="True" MulitselectPrompt="Multi allowed" EntrySeperator=" ">
            <Choice ChoiceID="1" ChoiceText="Choice 1" />
            <Choice ChoiceID="2" ChoiceText="Choice 2" />
            <Choice ChoiceID="3" ChoiceText="Choice 3" />
            <Choice ChoiceID="4" ChoiceText=" " />            
        </ChoiceInput> 
    </UserProvided>
</MyWordTemplateDefinition>
```
As you can see, there are this elements with an `Input`-postfix, like `UserInput`. These are the elements that define what the user has to provide, if she or he chooses to create a word template of the type `protocol`, provided by the definition file name AND the `Name` attribute of the document element `MyWordTemplateDefinition`.

When the input work is done the script will then load the Word template document, which it will find in the local folder `.\templates`, and which has the same name as the template definition file. In this example, this would be `.\templates\protocol.docx`. The script will then replace all the placeholders that match the IDs of the input elements in the template definitions file. For example, if there is a placeholder `test` in the Word template it will be replaced by the user input for the XML element `UserInput`. Of course, the script will not replace every occurance of `test` in the Word document. The placeholder has to be surrounded by `!@` and `@!`. So the correct form would be `!@test@!`.

## Template Definition Format
The correct syntax and semantics of the template definition files are tested both at loading time, as well as exectuion time. During loading of the XML file, the schema is validated against the schema definition file `.\templateDefinitions\TemplateDefinition.xsd`. Please check the definition yourself, if you are intrested in it. It is in the repository, of course.
## Template Definition Input Elements
There are currently only a few input elements implemented.
### UserInput
The `UserInput` element lets the user provide freetext or text validated against a regular expression. You can provide such a regular expression in the `ValidateRegex' attribute.
#### ConditionYoungerThan
You can provide date conditions, so that you can ensure, that some date is younger than another date. Use the `ConditionDateYoungerThan` attribute for that. The older date has to be definded before the younger date.
#### Examples
```
...
        <UserInput Name="Some Number" Prompt="Enter some number" ID="somenumber" ValidateRegex="^[0-9]+$" /> 
...
```

```
...
<UserInput Name="Some Older Date" Prompt="Older Date (z.B. 01.01.2019)" ID="olderdate" ValidateRegex="^[0-9]{2}.[0-2]?[0-9]{1}.[0-9]{4}$" />
...
<UserInput Name="Some Younger Date" Prompt="Younger Date (e.g. 01.01.2019)" ID="youngerdate"    ValidateRegex="^[0-9]{2}.[0-2]?[0-9]{1}.[0-9]{4}$" ConditionDateYoungerThan="olderdate" />
...
```
### ChoiceInput
You can provide choices to the user, so that she or he only has to select them by providing the according number. There are single and multi choices, that can be provided.
#### Single Choice Example
```
    <ChoiceInput Name="Asservatenart" Prompt="Art des Asservats" ID="asservatentyp" 
        BreakKeyword="cancle" FooterPrompt="Please choose" ErrorPrompt="Invalid"
        AllowMultiselect="False">
        <Choice ChoiceID="1" ChoiceText="An apple" />
        <Choice ChoiceID="2" ChoiceText="A pie" />
        <Choice ChoiceID="3" ChoiceText="En potet" />
        <Choice ChoiceID="4" ChoiceText="Flerer poteter" />
    </ChoiceInput>
```
#### Multi Choice Example
```
    <ChoiceInput Name="Asservatenart" Prompt="Art des Asservats" ID="asservatentyp" 
        BreakKeyword="cancle" FooterPrompt="Please choose" ErrorPrompt="Invalid"
        AllowMultiselect="True" MulitselectPrompt="Mehrfachauswahl mit Kommas separiert möglich."
        EntrySeperator=" ">
        <Choice ChoiceID="1" ChoiceText="An apple" />
        <Choice ChoiceID="2" ChoiceText="A pie" />
        <Choice ChoiceID="3" ChoiceText="En potet" />
        <Choice ChoiceID="4" ChoiceText="Flerer poteter" />
    </ChoiceInput>
```
The `EntrySeperator` in the above example will be used to seperate the user selections in the Word document later. In the example every selected `ChoiceText` will be entered into the Word document, seperated by a blank. You can use the value `NEWLINE` as `EntrySeperator`. In that case every entry will be seperated by new line in the Word document.
### LoopInput
You can also ask the user for more than one value, but always with the same general idea for the specific value. For example, if you want to enter different car brands, and every brand should be entered in the same spot in the Word template seperated by one another. In that case you can use the `LoopInput` elmenet. `LoopInput` uses nested `UserInput` elements for the actual input provision mechanism.
#### Example
```
    <LoopInput Name="Car Brands" ID="carbrand" BreakKeyword="done" Prompt="Provide car brands">
        <UserInput Name="Name of the brand" Prompt="Name of the brand" ID="brandname" />
        <UserInput Name="Country of the brand" Prompt="Country of the brand" ID="brandcountry" />
    </LoopInput>
```
### Placeholder
For simple text replacement you can use the `Placeholder` element. **Currently**, this only supports text, that then will be inserted at the according spot in the Word template document.
#### Example
```
    <Placeholders>
        <Placeholder Name="A placeholder" ID="aplaceholder">Some text</Placeholder>
    <Placeholders>
```
