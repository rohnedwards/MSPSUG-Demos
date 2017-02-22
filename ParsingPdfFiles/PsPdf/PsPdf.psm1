
Add-Type -ReferencedAssemblies $PSScriptRoot\itextsharp.dll @"
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

function RotateRect {
<#
PDF pages can be rotated. When that happens, you want the original
bounding rectangle's coordinates to be rotated in a way that keeps
the origin in the lower left corner so that the sorting and grouping
logic for all PDFs can stay the same.

180 and 270 (or -90) rotations are very simple, but I'm waiting on
a test PDF with that rotation before implementing them to make sure
I'm not overlooking something.
#>
    param(
        # Assumes clockwise rotation about origin
        [int] $Rotation,
        [System.Windows.Rect] $Rect,
        [int] $PageWidth,
        [int] $PageHeight
    )

    switch ($Rotation) {
        0 { $Rect }
        90 { 
            $NewX = $Rect.Y
            $NewY = $PageHeight + ($Rect.X * -1)
            [System.Windows.Rect]::new($NewX, $NewY, $Rect.Height, $Rect.Width) 
        }
        default {
            throw "Unsupported rotation!"
        }
    }
}

function Get-PdfTextChunk {
<#
Bounding rectangles have the property that the bottom/top are backwards due to
the origin of the PDF being in the lower left corner.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.IO.FileInfo] $InputObject,
        [int] $StartPage = 1,
        [int] $EndPage
    )

    process {
        
        if (-not (Test-Path $InputObject.FullName)) {
            Write-Error "Invalid file path: $($InputObject.FullName)"
            return
        }

        $Reader = [iTextSharp.text.pdf.PdfReader]::new($InputObject.FullName)
        try {

            $EndPage = if ($PSBoundParameters.ContainsKey('EndPage')) {
                if ($EndPage -gt $Reader.NumberOfPages) {
                    Write-Warning "-EndPage parameter (${EndPage}) is greater than the number of pages in the PDF ($($Reader.NumberOfPages))"
                    $Reader.NumberOfPages
                }
                else {
                    $EndPage
                }
            }
            else {
                $Reader.NumberOfPages
            }

            for ($i = 1; $i -le $EndPage; $i++) {

                $CurrentPage = $i
                # Not interested in the rendered text (the TextRenderInfo objects will have all that info)
                $ExtractionStrategy = [ROE.RawRenderInfoExtractionStrategy]::new()

                [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($Reader, $i, $ExtractionStrategy) | Out-Null

                $ExtractionStrategy.GetTextRenderInfoList() | ForEach-Object {
                    
                    $CurrentObject = $_

                    $ascentstart = $CurrentObject.GetAscentLine().GetStartPoint()
                    $descentend = $CurrentObject.GetDescentLine().GetEndPoint()
                    $br = [System.Windows.Rect]::new(
                        [System.Windows.Point]::new([decimal]::Round($ascentstart[0], 2), [decimal]::Round($ascentstart[1], 2)),
                        [System.Windows.Point]::new([decimal]::Round($descentend[0]), [decimal]::Round($descentend[1]))
                    )

                    $PageSize = $Reader.GetPageSizeWithRotation($CurrentPage)
                    $br = RotateRect -Rotation $PageSize.Rotation -Rect $br -PageWidth $PageSize.Width -PageHeight $PageSize.Height

                    [PSCustomObject] @{
                        Page = $CurrentPage
                        Text = $CurrentObject.GetText()
                        Font = $CurrentObject.GetFont().PostscriptFontName
                        SpaceWidth = $CurrentObject.GetSingleSpaceWidth()
                        BoundingRect = $br
                    }
                }
            }
        }
        catch {
            Write-Error "Error working with '$($InputObject.FullName)' PDF file: $_"
        }
        finally {
            $Reader.Dispose()
        }
    }
}

function EatRectangle {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $InputRectangle,
        [Parameter(Mandatory, Position=0)]
        $Food
    )

    process {

        # Rectangles have bottom/top reversed. That shouldn't matter here, though,
        # since we're just dealing with bounding rectangles that are already reversed.
        # We shouldn't care about the fact that they're switched until we need to
        # compare positions

        $Points = $InputRectangle, $Food
        $Top = $Points | sort Top | select -ExpandProperty Top -First 1
        $Left = $Points | sort Left | select -ExpandProperty Left -First 1
        $Right = $Points | sort Right | select -ExpandProperty Right -Last 1
        $Bottom = $Points | sort Bottom | select -ExpandProperty Bottom -Last 1

        $TopLeft = [System.Windows.Point]::new($Left, $Top)
        $BottomRight = [System.Windows.Point]::new($Right, $Bottom)

        [System.Windows.Rect]::new($TopLeft, $BottomRight)
    }
}

function Get-PdfRectangle {
<#

This takes a set of rectangles, and generates either row rectangles or
column rectangles that have been merged based on the -GapSize.

#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IEnumerable] $InputRectangles,
        # Vertical gaps between text chunks greater than this number will cause new rectangles
        # to be created. Defaults to max int size, which means columns won't be split up
        [double] $GapSize = [int]::MaxValue,
        [ValidateSet('Row', 'Column')]
        [string] $Mode = 'Row'
    )

    <#
        Whether you're looking for rectangles in Row mode (that's left to right) or
        Column mode (top to bottom), the logic is basically the same. There are some
        differences, though:
        
        GapSize:
            This is the size of the gap b/w the previous and current rectangle in the
            same row or column. If in row mode, this is the difference b/w the right
            edge of the previous node and the left edge of the current node. In column
            mode, it's the difference b/w the bottom edge of the previous (which is
            actually the 'Top' property) and the top edge of the current (which is the
            'Bottom' property)


        MidpointCalculation:
            This is used for grouping rectangles together to find out if they belong
            in the same row or column. In row mode, we try to group all rectangles by
            their top edge ('Bottom' property), then check each grouping for whether
            or not they should be merged, then we group again by the bottom edge ('Top'
            property), then by the midpoint. Columns go by the Left, Right, then
            midpoint. Rectangles don't have a midpoint property, so we create a lambda
            scriptblock to handle that, and use it during sorting and grouping.

        GroupAndSortProps:
            Explained in previous section. This is just where we assign the hashtable
            that will be enumerated later

        UngroupedChunkBeforeMe:
            Imagine figuring out column groups where the top and bottom of the page
            have lines that line up perfectly on their left edge, and there's a table
            b/w them that is further to the right. If you use the normal grouping logic,
            you'd end up with columns that overlap. This lambda scriptblock has logic
            to figure out if there's some chunk of text b/w the previous and current that
            wasn't captured in the current group, and if so, it won't merge those rectangles.
    
    
        Of course, all of this may be parameterized later...
    #>
    switch ($Mode) {
        Row {
            $GetGapSize = {
                $Current.Left - $Previous.Right
            }

            $MidpointCalculation = {
                [math]::Round($_.Height/2 + $_.Top, 0)
            }

            # First, sort by Top, then bottom, then midpoint
            $GroupAndSortProps = @{
                # Don't forget, Bottom means we're grouping by the top
                Bottom = echo @{Expression = 'Bottom'; Descending = $true}, Left
                Top = echo @{Expression = 'Top'; Descending = $true}, Left
                $MidpointCalculation = echo $MidpointCalculation, Left
            }

            $UngroupedChunkBeforeMe = {
                # Row mode doesn't support this yet
                $false
            }
        }

        Column {
            $GetGapSize = {
                # This says to take the bottom of the previous rect and subtract it from the top
                # of the current
                $Previous.Top - $Current.Bottom
            }

            $MidpointCalculation = {
                [math]::Round($_.Width/2 + $_.Left, 0)
            }

            $GroupAndSortProps = @{
                Left = echo Left, @{Expression = 'Bottom'; Descending = $true}
                Right = echo Right, @{Expression = 'Bottom'; Descending = $true}
                $MidpointCalculation = echo $MidpointCalculation, @{Expression = 'Bottom'; Descending = $true}
            }

            $UngroupedChunkBeforeMe = {
                [bool] ($Grouped.Group | where {
                    $_.Bottom -lt $Previous.Top -and $_.Bottom -gt $Current.Bottom
                })
            }
        }

        default {
            throw "Unknown Mode: ${Mode}"
        }
    }

    foreach ($CurrentGroupSort in $GroupAndSortProps.GetEnumerator()) {
    
        Write-Verbose "InputRectangles.Count = $($InputRectangles.Count)"
        Write-Verbose "  ->  Sorting on: $($CurrentGroupSort.Value)"
        Write-Verbose "  -> Grouping on: $($CurrentGroupSort.Key)"
        Write-Debug "GroupSort: $($CurrentGroupSort.Key)"
        $Grouped = $InputRectangles | sort $CurrentGroupSort.Value | group $CurrentGroupSort.Key
        $InputRectangles = foreach ($Group in $Grouped) {
        
            $NewRect = $null

            Write-Verbose "  -> New group: $($CurrentGroupSort.Key) = $($Group.Name)"
            for ($i = 0; $i -lt $Group.Count; $i++) {
                $Current = $Group.Group[$i]
                $Previous = if ($i -gt 0) {
                    $Group.Group[$i - 1]
                }

                $GapSizeFromPrevious = if ($Previous) {
                    & $GetGapSize
                }
                else {
                    [int]::MinValue
                }            

                if ($i -gt 0 -and $GapSizeFromPrevious -gt $GapSize) {
                    Write-Verbose "  -> GapSizeFromPrevious = ${GapSizeFromPrevious}; creating new rectangle"
                    # Found a gap large enough to create a new rectangle. Dump the previous one and clear it out
                    $NewRect
                    $NewRect = $null
                }
                elseif (& $UngroupedChunkBeforeMe) {
                    Write-Verbose '  -> Ungrouped chunk before this rect detected, so creating new rectangle'
                    $NewRect
                    $NewRect = $null
                }

                $NewRect = if ($null -eq $NewRect) {
                    [System.Windows.Rect]::new($Current.TopLeft, $Current.BottomRight)
                }
                else {
                    Write-Verbose "  -> Combining (${Current}) with (${NewRect})"
                    $Current | EatRectangle -Food $NewRect -Verbose:$false
                }
            }

            # Output final column rect:
            if ($null -ne $NewRect) {
                $NewRect
            }
        }
    }

    $InputRectangles
}

function Get-PdfTextFromRectanglesHelper {
    [CmdletBinding()]
    param(
        $TextChunks,
        [System.Windows.Rect[]] $RowRectangles,
        [System.Windows.Rect[]] $ColumnRectangles
    )

    # We don't want to change the objects that are passed:
    $TextChunks = $TextChunks | ForEach-Object {
        $_.psobject.Copy()
    }

    #region Tag each text chunk with its row and column

    # Sort rows from top to bottom (remember that b/c of Rect class
    # and PDF origin differences, the Bottom property is the visual
    # top edge
    $RowRectangles = $RowRectangles | sort Bottom -Descending
    for ($i = 0; $i -lt $RowRectangles.Count; $i++) {
        $Rectangle = $RowRectangles[$i]

        for ($j = 0; $j -lt $TextChunks.Count; $j++) {

            $CurrentChunk = $TextChunks[$j]

            if ($Rectangle.IntersectsWith($CurrentChunk.BoundingRect)) {
                $TextChunks[$j] | Add-Member -NotePropertyName Row -NotePropertyValue $i -Force
            }
        }
    }

    $ColumnRectangles = $ColumnRectangles | sort Left
    for ($i = 0; $i -lt $ColumnRectangles.Count; $i++) {
        $Rectangle = $ColumnRectangles[$i]

        for ($j = 0; $j -lt $TextChunks.Count; $j++) {

            $CurrentChunk = $TextChunks[$j]

            if ($Rectangle.IntersectsWith($CurrentChunk.BoundingRect)) {
                $TextChunks[$j] | Add-Member -NotePropertyName Column -NotePropertyValue $i -Force
            }
        }
    }
    #endregion


    $ReturnObj = $null
    $LastColumn = $null

    $WriteReturnObj = {
        $Type = switch ($ReturnObj.GetType().Name) {
            ArrayList { 
                'Table' 
Write-Warning "Don't forget to do table column intersect check"
            }
            StringBuilder { 
                'TextBlock' 
                $ReturnObj = $ReturnObj.ToString() | foreach Trim
            }
            default { 'Unknown' }
        }

        [PSCustomObject] @{
            Type = $Type
            Output = $ReturnObj
        }

        # Clear it out so the next one can be built
        Set-Variable -Name ReturnObj -Value $null -Scope 1
    }

    $TextChunks | group Row | ForEach-Object {
        $RowGroup = $_.Group
        $ColumnGroup = @($RowGroup | group Column)
        $IsTable = $ColumnGroup.Count -gt 1

        if ($IsTable) {
            if ($null -ne $ReturnObj -and @($LastColumn).Count -ne @($ColumnGroup).Count) {

                # $ReturnObj has something, and it doesn't match our current table info, so
                # write it out (which will also empty it)
                & $WriteReturnObj

                # This is an attempt to give each table that's detected its own 'Column0', 'Column1', 'Column[n-1]'
                # set of columns (and it helps when columns didn't line up exactly right)
                $LastColumn = for ($i = 0; $i -lt $ColumnGroup.Count; $i++) {
                    $ColumnGroup[$i].Group[0] | select Left, Right | Add-Member -NotePropertyName ColumnName -NotePropertyValue "Column${i}" -PassThru
                }
            }

            # Tables are arraylists of pscustomobjects    
            $ReturnObjProperties = [ordered] @{}
            for ($z = 0; $z -lt $LastColumn.Count; $z++) {
                $ReturnObjProperties[$LastColumn[$z].ColumnName] = $ColumnGroup[$z].Group.Text -join ''
            }

            if ($ReturnObj -isnot [System.Collections.ArrayList]) {
                # We need to create an arraylist
                $ReturnObj = [System.Collections.ArrayList]::new()
            }

            $ReturnObj.Add([PSCustomObject] $ReturnObjProperties) | Out-Null
        }
        else {
            if ($null -ne $ReturnObj -and $LastColumn -ne $_.Group[0].Column) {
                & $WriteReturnObj

            }


            if ($ReturnObj -isnot [System.Text.StringBuilder]) {
                # We need to create a new stringbuilder
                $ReturnObj = [System.Text.StringBuilder]::new()
            }

            $LastColumn = $_.Group[0].Column
            $Value = $_.Group.Text -join ''
            
            $ReturnObj.AppendLine($Value) | Out-Null
        }
    }

    # Get rid of the $ReturnObject that's waiting
    & $WriteReturnObj
}

function Get-PdfText {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string] $FilePath
    )

    Get-PdfTextChunk $FilePath -ErrorAction Stop | group Page | ForEach-Object {

        $AverageWidth = $_.Group | measure -Property SpaceWidth -Average | select -ExpandProperty Average
        $WholeRows = Get-PdfRectangle $_.Group.BoundingRect #-GapSize $AverageWidth
        $WholeColumns = Get-PdfRectangle $WholeRows -Mode Column #-GapSize ($AverageWidth/2)

        Get-PdfTextFromRectanglesHelper -TextChunks $_.Group -RowRectangles $WholeRows -ColumnRectangles $WholeColumns | Add-Member -NotePropertyMembers @{
            Path = $FilePath
            Page = $_.Name
        } -PassThru | ForEach-Object {
            if ($_.Type -eq 'Table') {
                Add-Member -InputObject $_.Output -MemberType ScriptMethod -Name ToString -Value {
                    "tostringed!"
                } -Force
            }
            $_
        }
    }
}

function New-PdfRectangleCanvas {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string] $FilePath,
        [System.Collections.IDictionary] $ColorsAndRectangles = @{},
        [System.Collections.IDictionary] $ColorsAndTextChunks = @{},
        [double] $ResizeFactor = 1
    )

    try {
        $Pdf = [iTextSharp.text.pdf.PdfReader]::new($FilePath)
    }
    catch {
        Write-Warning "Error creating PdfReader: ${_}"
        return
    }

    $PageSize = $Pdf.GetPageSizeWithRotation(1)

    [int] $PageHeight = $PageSize.Height * $ResizeFactor
    [int] $PageWidth = $PageSize.Width * $ResizeFactor

    [xml] $Xaml = @"
    <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="Initial Window" WindowStartupLocation = "CenterScreen"
        Width = "$($PageWidth)" Height = "$($PageHeight)" ShowInTaskbar = "True" ResizeMode = "NoResize" >

            <Border x:Name="CanvasBorder" BorderBrush="Black" BorderThickness = "1" >
                <Canvas x:Name="Canvas" />
            </Border>
    </Window>
"@

    try { 
        $Window = [Windows.Markup.XamlReader]::Load( [System.Xml.XmlNodeReader]::new($Xaml) )
        $Canvas = $Window.FindName('Canvas')
    }
    catch {
        Write-Warning "Error creating Window: ${_}"
        return
    }

    foreach ($ColorRect in $ColorsAndRectangles.GetEnumerator()) {
        $ColorBrush = [System.Windows.Media.SolidColorBrush]::new($ColorRect.Key)

        foreach ($BR in $ColorRect.Value) {
            $Rectangle = [System.Windows.Shapes.Rectangle]::new()
            $Rectangle.Width = $BR.Width * $ResizeFactor
            $Rectangle.Height = $BR.Height * $ResizeFactor

            $BrLeft = $BR.Left * $ResizeFactor
            $BrTop = $BR.Top * $ResizeFactor
            $BrBottom = $BR.Bottom * $ResizeFactor
            $Rectangle.Stroke = $ColorBrush
            [System.Windows.Controls.Canvas]::SetLeft($Rectangle, $BrLeft)
            [System.Windows.Controls.Canvas]::SetTop($Rectangle, $PageHeight - $BrBottom)
            $Canvas.Children.Add($Rectangle) | Out-Null
        }
    }

    foreach ($ColorTextBlob in $ColorsAndTextChunks.GetEnumerator()) {
        $ColorBrush = [System.Windows.Media.SolidColorBrush]::new($ColorTextBlob.Key)

        foreach ($TB in $ColorTextBlob.Value) {
            $ViewBox = [System.Windows.Controls.Viewbox]::new()
            $ViewBox.Width = $TB.BoundingRect.Width * $ResizeFactor
            $ViewBox.Height = $TB.BoundingRect.Height * $ResizeFactor
            $ViewBox.Stretch = 'Uniform'
            $ViewBox.StretchDirection = 'DownOnly'

            $TextBlock = [System.Windows.Controls.TextBlock]::new()
            $TextBlock.Foreground = $ColorBrush
            $TextBlock.Text = $Tb.Text

try {
$TextBlock.FontFamily = [System.Windows.Media.FontFamily]::new($TB.Font)
}
catch {
    Write-Warning "Couldn't create FontFamily for '$($TB.Font)': ${_}"
}

            $BrLeft = $TB.BoundingRect.Left * $ResizeFactor
            $BrTop = $TB.BoundingRect.Top * $ResizeFactor
            $BrBottom = $TB.BoundingRect.Bottom * $ResizeFactor

            $ViewBox.AddChild($TextBlock)


            [System.Windows.Controls.Canvas]::SetLeft($ViewBox, $BrLeft)
            [System.Windows.Controls.Canvas]::SetTop($ViewBox, $PageHeight - $BrBottom)
            $Canvas.Children.Add($ViewBox) | Out-Null
        }
    }


    $Window.ShowDialog()
    $Window.Close()
    $Pdf.Dispose()
}