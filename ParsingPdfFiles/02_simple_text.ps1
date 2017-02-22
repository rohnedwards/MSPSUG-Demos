$PresentationPath = $PWD.Path

# iTextSharp library can be found here: https://github.com/itext/itextsharp/releases/download/5.5.10/itextsharp-all-5.5.10.zip
Add-Type -Path "$PresentationPath\itextsharp.dll"

break

#region Notes
#
# So pulling form control names and values is relatively simple. What about getting the raw text? Well,
# that's a little harder to do. It turns out you can't simply ask a PDF file to tell you what text it
# contains. That's because a PDF file has instructions for things to print, and you, as the person viewing
# the file, can't be sure what order anything happens in. PDF readers execute all of the print instructions,
# and then show you what the final product looks like.
#
# PDF helper classes like iTextSharp know how to render this information, and they offer ways to try to
# extract this text for you. In iText land, you create an instance of a TextExtractionStrategy object. iText
# comes with two of these: SimpleTextExtractionStrategy and LocationTextExtractionStrategy.
#
#endregion

#region SimpleTextExtractionStrategy example
# Let's start with the simple strategy, using the PDF from the last example:
$PdfPath = "$PresentationPath\sample_pdfs\interactiveform_enabled.pdf"
& $PdfPath

$PdfReader = [iTextSharp.text.pdf.PdfReader]::new($PdfPath)

$PageNumber = 1
$ExtractionStrategy = [iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy]::new()
[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PdfReader, $PageNumber, $ExtractionStrategy) | 
    Tee-Object -Variable SteStrat

#endregion

#region LocationTextExtractionStrategy example
# Let's see if the other default extraction strategy that comes with iTextSharp, the location strategy,
# is any different:
$PageNumber = 1
$ExtractionStrategy = [iTextSharp.text.pdf.parser.LocationTextExtractionStrategy]::new()
[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PdfReader, $PageNumber, $ExtractionStrategy) | 
    Tee-Object -Variable LteText

# In this case, the location based extraction strategy looks better. That seems to usually
# be the case, and is probably why GetTextFromPage() without an extraction strategy provided
# defaults to using the location strategy.

$PdfReader.Close()
#endregion

#region Another example with a different PDF
# PDF link: http://www.orimi.com/pdf-test.pdf
$PdfPath = "$PresentationPath\sample_pdfs\pdf-test.pdf"
& $PdfPath

$PdfReader = [iTextSharp.text.pdf.PdfReader]::new($PdfPath)
[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PdfReader, 1, [iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy]::new()) | Tee-Object -Variable SteText
[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PdfReader, 1, [iTextSharp.text.pdf.parser.LocationTextExtractionStrategy]::new()) | Tee-Object -Variable LteText
$PdfReader.Close()

# Because of the simplicity of this one's content, the extraction strategies give the same results.
#endregion

#region Invoice example
# PDF link: http://www.princexml.com/howcome/2016/samples/invoice/index.pdf
$PdfPath = "$PresentationPath\sample_pdfs\simpleinvoicesample.pdf"
& $PdfPath
$PdfReader = [iTextSharp.text.pdf.PdfReader]::new($PdfPath)

$CurrentPage = 1
[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PdfReader, $CurrentPage, [iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy]::new()) |
    Tee-Object -Variable SteInvoice

[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PdfReader, $CurrentPage, [iTextSharp.text.pdf.parser.LocationTextExtractionStrategy]::new()) |
    Tee-Object -Variable LteInvoice

$PdfReader.Close()
#endregion

# So, location and simple text extraction strategies each have their own advantages and
# disadvantages, and which one's better really depends on the PDF, and the user's ability
# to verify that the information is accurate. That doesn't sound like automation...
#
# Thankfully, iText allows you create your own extraction strategies to try to handle the
# shortcomings of the default strategies. See the next demo file for an extraction 
# strategy that offloads all of the hard work to PowerShell for figuring out what text
# goes where.
