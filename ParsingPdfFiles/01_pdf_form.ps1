$PresentationPath = $PWD.Path

# iTextSharp library can be found here: https://github.com/itext/itextsharp/releases/download/5.5.10/itextsharp-all-5.5.10.zip
Add-Type -Path "$PresentationPath\itextsharp.dll"

break

#region Example of reading from PDF form
# PDF source: http://help.adobe.com/en_US/Acrobat/9.0/Samples/interactiveform_enabled.pdf
$PdfPath = "$PresentationPath\sample_pdfs\interactiveform_enabled.pdf"
& $PdfPath

# Let's look at the form fields using iTextSharp:
# First, create a new PdfReader instance:
$PdfReader = [iTextSharp.text.pdf.PdfReader]::new($PdfPath)

# Fields are found in the AcroFields.Fields property:
$PdfReader.AcroFields.Fields

# AcroFields contains helper methods for working with the fields
$PdfReader.AcroFields | Get-Member 
$PdfReader.AcroFields.GetField('Name_Last')  # Empty b/c no value textbox
$PdfReader.AcroFields.GetFieldType('Name_Last')  # Empty b/c no value textbox
$PdfReader.AcroFields.GetFieldType('Print')

# GetFieldType() returns a number. That number can be found by looking for static
# FIELD_TYPE_* members on the AcroFields class; 1 is FIELD_TYPE_PUSHBUTTON
[iTextSharp.text.pdf.AcroFields]::FIELD_TYPE_PUSHBUTTON
[iTextSharp.text.pdf.AcroFields]::FIELD_TYPE_TEXT

# Make a hashtable of PDF field types so we can easily look them up:
$FieldTypes = @{}
[iTextSharp.text.pdf.AcroFields] | Get-Member -Static FIELD_TYPE_* | ForEach-Object {
    $Name = $_.Name -replace '^FIELD_TYPE_'
    $Key = ($_.TypeName -as [type])::($_.Name)
    $FieldTypes[$Key] = $Name
}

# Run this and notice all the values are blank since the form is empty
foreach ($CurrentFieldName in $PdfReader.AcroFields.Fields.Keys) {
    [PSCustomObject] @{
        Type = $FieldTypes[$PdfReader.AcroFields.GetFieldType($CurrentFieldName)]
        Name = $CurrentFieldName
        Value = $PdfReader.AcroFields.GetField($CurrentFieldName)
    }
}

# If you fill in some values in the form, save it, then re-open it, you should
# see the values when run the foreach loop above
$PdfReader.Close()
$PdfPath = "$PresentationPath\sample_pdfs\interactiveform_enabled_filled.pdf"
$PdfReader = [iTextSharp.text.pdf.PdfReader]::new($PdfPath)
# Now go back and run the foreach() loop above

# This could very easily be converted into a re-usable tool.

# Notice that the names aren't exactly the same as the form's nice labels. That's because,
# just like with Windows forms and WPF programming, the name of the underlying field is
# different than the text that is used on the form. When looking at this form, you'll notice
# that there's a 'NAME' section, and inside that is the 'Last' textbox. The field name for
# that is 'Name_Last', though.

#endregion