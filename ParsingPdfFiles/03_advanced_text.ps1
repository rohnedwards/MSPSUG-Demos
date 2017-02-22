$PresentationPath = $PWD.Path

# iTextSharp library can be found here: https://github.com/itext/itextsharp/releases/download/5.5.10/itextsharp-all-5.5.10.zip
Add-Type -Path "$PresentationPath\itextsharp.dll"

break

#region A PowerShell extraction strategy
# The beaty of iTextSharp is that it provides enough tools for you to create your own text extraction
# strategies. If you used C#, you'd create classes that implemented the ITextExtractionStrategy
# interface, and you'd come up with some strategies and algorithms to output a string that represented
# a page of text from a PDF.
#
# See below for a minimal implementation of the interface, that gives you access to a list of the
# raw rendering objects that iTextSharp uses:

Add-Type -ReferencedAssemblies $PresentationPath\itextsharp.dll @"
using System;
using System.Text;
using System.Collections.Generic;
using iTextSharp.text.pdf.parser;

namespace ROE
{
	public class RawRenderInfoExtractionStrategy : ITextExtractionStrategy, IRenderListener
	{
        public RawRenderInfoExtractionStrategy() { }
        private List<TextRenderInfo> textRenderInfoList = new List<TextRenderInfo>();
        private List<ImageRenderInfo> imageRenderInfoList = new List<ImageRenderInfo>();

        public List<TextRenderInfo> GetTextRenderInfoList() {
            return textRenderInfoList;
        }

        public List<ImageRenderInfo> GetImageRenderInfoList() {
            return imageRenderInfoList;
        }

		public virtual void BeginTextBlock() { }
		public virtual void EndTextBlock() { }
		public virtual string GetResultantText() {
			return "No resultant text. Call GetTextRenderInfoList() to get raw render info objects";
		}
		public virtual void RenderText(TextRenderInfo renderInfo) {
            textRenderInfoList.Add(renderInfo);			
		}
		public virtual void RenderImage(ImageRenderInfo renderInfo) {
            imageRenderInfoList.Add(renderInfo);			
		}
	}
}
"@
#endregion

#region Revist pdf-test
$PdfPath = "$PresentationPath\sample_pdfs\pdf-test.pdf"
& $PdfPath

$PdfReader = [iTextSharp.text.pdf.PdfReader]::new($PdfPath)
$RawRenderInfoStrategy = [ROE.RawRenderInfoExtractionStrategy]::new()
$CurrentPage = 1
[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PdfReader, $CurrentPage, $RawRenderInfoStrategy)

# After calling that, $RawRenderInfoStrategy.GetTextRenderInfo
$TextChunks = $RawRenderInfoStrategy.GetTextRenderInfoList() | ForEach-Object {

    $InputObject = $_

    $ascentstart = $InputObject.GetAscentLine().GetStartPoint()
    $descentend = $InputObject.GetDescentLine().GetEndPoint()
    $br = [System.Windows.Rect]::new(
        [System.Windows.Point]::new([decimal]::Round($ascentstart[0], 2), [decimal]::Round($ascentstart[1], 2)),
        [System.Windows.Point]::new([decimal]::Round($descentend[0]), [decimal]::Round($descentend[1]))
    )

    # NOTE: This method has a huge design flaw: it doesn't take page rotation into account. If rotation has been
    # applied to the page, you'll want to change the coordinates so the origin is always at the bottom left so
    # that any position logic you use will work no matter rotation

    $PageSize = $PdfReader.GetPageSizeWithRotation($CurrentPage)
    if ($PageSize.Rotation) {
        throw 'Page has rotation applied, so can''t figure out rectangles'
    }

    [PSCustomObject] @{
        Page = $CurrentPage
        Text = $InputObject.GetText()
        Font = $InputObject.GetFont().PostscriptFontName
        SpaceWidth = $InputObject.GetSingleSpaceWidth()
        BoundingRect = $br
    }
}
$TextChunks | ft

#region Notes
#
# Whoa! some of the text that was on the same lines is now split up. What's going on here?
# What you see when a PDF is rendered isn't necessarily how the lines (or chunks as iText
# calls them) were printed. You're seeing a much more raw view of how this particular PDF
# is organized
#
#endregion

# This is a helper command from the PsPdf module that we'll use later...
# The text that's printed here isn't to scale, so it's going to look a little odd...
Import-Module "$PresentationPath\PsPdf"
$PdfPath | New-PdfRectangleCanvas -ResizeFactor .8 -ColorsAndRectangles @{
    Red = $TextChunks.BoundingRect
} -ColorsAndTextChunks @{
    Black = $TextChunks
}

$PdfReader.Close()

#region Combine the text
# From here on out, we're going to use Get-PdfTextChunk from the PsPdf module
# to handle opening a PdfReader and parsing through the TextRenderInfo objects.
# It does the same as we did in the ForEach-Object block in the last example.
#
$TextChunks = $PdfPath | Get-PdfTextChunk | where Page -eq 1 #| where { -not ($_.Text -match '^\s*$') }

# We need a way to figure out which rectangles are on the same lines and are close enough that they
# should be joined together:
$RectanglesWithText = $TextChunks.ForEach{$_.BoundingRect | Add-Member -NotePropertyName Text -NotePropertyValue $_.Text -PassThru}
$RectanglesWithText | ft
$RectanglesWithText | sort @{E='Bottom'; Descending=$true}, Left | ft -GroupBy Bottom -Property Text, Left, Bottom, Right, Top

# Determine the average horizontal space width:
$AverageSpaceWidth = $TextChunks | Measure-Object -Property SpaceWidth -Average -Maximum -Minimum | select -ExpandProperty Average

# Get-PdfRectangle basically does the work we just did, except it tries to group on more than just one edge, and it
# can do the same work for columns.
$CombinedRows = Get-PdfRectangle -InputRectangles $TextChunks.BoundingRect -GapSize $AverageSpaceWidth
$WholeRows = Get-PdfRectangle -InputRectangles $TextChunks.BoundingRect

$PdfPath | New-PdfRectangleCanvas -ResizeFactor .8 -ColorsAndRectangles @{
    #Red = $TextChunks.BoundingRect
    Green = $CombinedRows
} -ColorsAndTextChunks @{
    Black = $TextChunks
}

$WholeRows = $WholeRows | sort Bottom -Descending
$LineSpaceSizes = for ($i = 1; $i -lt $WholeRows.Count; $i++) {
    $Previous = $WholeRows[$i - 1]
    $Current = $WholeRows[$i]

    $Previous.Top - $Current.Bottom
}

$AverageLineSpace = $LineSpaceSizes | where { $_ -gt 0 } | measure -Average | select -ExpandProperty Average

$CombinedColumns = Get-PdfRectangle $CombinedRows -Mode Column -GapSize $AverageLineSpace
$WholeColumns = Get-PdfRectangle $WholeRows -Mode Column

$PdfPath | New-PdfRectangleCanvas -ResizeFactor .8 -ColorsAndRectangles @{
    #Red = $TextChunks.BoundingRect
    Green = $CombinedRows
    Blue = $CombinedColumns
} -ColorsAndTextChunks @{
    Black = $TextChunks
}

Get-PdfTextFromRectanglesHelper -TextChunks $TextChunks -RowRectangles $CombinedRows -ColumnRectangles $CombinedColumns | 
    fl Type, @{N='Output'; E={ $_.Output | ft | Out-String | % Trim }}

Get-PdfTextFromRectanglesHelper -TextChunks $TextChunks -RowRectangles $WholeRows -ColumnRectangles $WholeColumns | 
    fl Type, @{N='Output'; E={ $_.Output | ft | Out-String | % Trim }}

#endregion

#endregion


#region Another look at the invoice
$PdfPath = "$PresentationPath\sample_pdfs\simpleinvoicesample.pdf"
& $PdfPath

$TextChunks = $PdfPath | Get-PdfTextChunk | where Page -eq 1 #| where { -not ($_.Text -match '^\s*$') }

# What do the raw rectangles look like?
$PdfPath | New-PdfRectangleCanvas -ColorsAndRectangles @{
    Red = $TextChunks.BoundingRect
} -ColorsAndTextChunks @{
    Black = $TextChunks
} -ResizeFactor .8

$AverageSpaceWidth = $TextChunks | measure SpaceWidth -Average | select -ExpandProperty Average
$CombinedRows = Get-PdfRectangle $TextChunks.BoundingRect -GapSize $AverageSpaceWidth
$WholeRows = Get-PdfRectangle $TextChunks.BoundingRect  # WholeRows would cause a fake table to be generated in this case
$CombinedRows2 = Get-PdfRectangle $CombinedRows -GapSize 200

$PdfPath | New-PdfRectangleCanvas -ColorsAndRectangles @{
    Red = $CombinedRows
} -ColorsAndTextChunks @{
    Black = $TextChunks
} -ResizeFactor .8

$CombinedColumns = Get-PdfRectangle $CombinedRows -Mode Column -GapSize ($AverageSpaceWidth * 3)
$PdfPath | New-PdfRectangleCanvas -ColorsAndRectangles @{
    Red = $CombinedRows2
    Green = $CombinedColumns
} -ColorsAndTextChunks @{
    Black = $TextChunks
} -ResizeFactor .8

Get-PdfTextFromRectanglesHelper -TextChunks $TextChunks -RowRectangles $CombinedRows2 -ColumnRectangles $CombinedColumns | 
    fl Type, @{N='Output'; E={ $_.Output | ft | Out-String | % Trim }}

#endregion

#region A different invoice
# PDF Link: http://www.princexml.com/samples/invoice/invoicesample.pdf
$PdfPath = "$PresentationPath\sample_pdfs\invoicesample.pdf"
& $PdfPath

$TextChunks = $PdfPath | Get-PdfTextChunk | where Page -eq 1 #| where { -not ($_.Text -match '^\s*$') }

# What do the raw rectangles look like?
$PdfPath | New-PdfRectangleCanvas -ColorsAndRectangles @{
    Red = $TextChunks.BoundingRect
} -ColorsAndTextChunks @{
    Black = $TextChunks
} -ResizeFactor .8

$AverageSpaceWidth = $TextChunks | measure SpaceWidth -Average | select -ExpandProperty Average
$CombinedRows = Get-PdfRectangle $TextChunks.BoundingRect -GapSize $AverageSpaceWidth
$PdfPath | New-PdfRectangleCanvas -ColorsAndRectangles @{
    Red = $CombinedRows
} -ColorsAndTextChunks @{
    Black = $TextChunks
} -ResizeFactor .8

$CombinedColumns = Get-PdfRectangle $CombinedRows -Mode Column -GapSize $AverageSpaceWidth
$PdfPath | New-PdfRectangleCanvas -ColorsAndRectangles @{
    Green = $CombinedColumns
} -ColorsAndTextChunks @{
    Black = $TextChunks
} -ResizeFactor .8

$CombinedRows = Get-PdfRectangle $CombinedColumns -GapSize $AverageSpaceWidth
$PdfPath | New-PdfRectangleCanvas -ColorsAndRectangles @{
    Red = $CombinedRows
} -ColorsAndTextChunks @{
    Black = $TextChunks
} -ResizeFactor .8

# Notice that the Total and its amount aren't grouped. Their Bottom edge, Top edge, and midpoint
# location must not match up. Get-PdfRectangle obviously needs some more tweaking.


$BigRows = Get-PdfRectangle $CombinedRows -GapSize 200
$BigColumns = Get-PdfRectangle $CombinedRows -Mode Column
$PdfPath | New-PdfRectangleCanvas -ColorsAndRectangles @{
    Red = $BigRows
    Green = $BigColumns
} -ColorsAndTextChunks @{
    Black = $TextChunks
} -ResizeFactor .8

Get-PdfTextFromRectanglesHelper -TextChunks $TextChunks -RowRectangles $BigRows -ColumnRectangles $BigColumns | 
    fl Type, @{N='Output'; E={ $_.Output | ft | Out-String | % Trim }}

#endregion
