#Requires -Version 7.0
[CmdletBinding()]
param([string]$Action = 'render', [string]$RunToken, [Parameter(Mandatory)][string]$OutputPath)

# Uses only the installed MathType commands. No vendor source is embedded in this test.
$root = Split-Path -Parent $PSScriptRoot
$bridgePath = Join-Path $root 'scripts\mathtype-word.ps1'
$source = Get-Content -Raw -LiteralPath $bridgePath
$begin = $source.IndexOf('Set-StrictMode -Version Latest')
$end = $source.LastIndexOf("`ntry {")
if ($begin -lt 0 -or $end -lt $begin) { throw 'Could not load the bridge helpers.' }
$helpers = $source.Substring($begin, $end - $begin).Replace('$PSScriptRoot', "'" + (Join-Path $root 'scripts').Replace("'", "''") + "'")
. ([scriptblock]::Create($helpers))
$script:PluginRoot = $root
$script:PackagedDefaultsPath = Join-Path $root 'config\defaults.json'

# Reuse conversion/placement, but let MathType render its own complete number format.
$ast = [Management.Automation.Language.Parser]::ParseFile($bridgePath, [ref]$null, [ref]$null)
$convert = $ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Convert-Equation'}, $true).Extent.Text
$convert = $convert.Replace('Normalize-SimpleNumberField -NumberField $numberField', '# Keep the native MathType chapter/section components.')
$convert = $convert.Replace('Normalized native number for', 'Preserved native number for')
$convert = $convert.Replace('to simple parenthesized format.', 'without simple-number normalization.')
. ([scriptblock]::Create($convert))

function Get-NativeSnapshot {
    $numbers = @(); $refs = @(); $breaks = @()
    foreach ($field in $script:Document.Fields) {
        $code = [string]$field.Code.Text
        $visible = $field.Code.Duplicate
        $visible.TextRetrievalMode.IncludeFieldCodes = $false
        $visible.TextRetrievalMode.IncludeHiddenText = $false
        $display = [regex]::Match([string]$visible.Text, '\([0-9.-]+\)').Value
        Release-ComObject $visible
        $row = [ordered]@{ text = $display; code = $code }
        if ($code -match '^\s*MACROBUTTON\s+MTPlaceRef\b') { $numbers += $row }
        elseif ($code -match '^\s*GOTOBUTTON\s+ZEqnNum\d+\b') {
            $target = [regex]::Match($code, '^\s*GOTOBUTTON\s+(ZEqnNum\d+)').Groups[1].Value
            $row['bookmark'] = $target
            $row['target_exists'] = [bool]$script:Document.Bookmarks.Exists($target)
            $refs += $row
        }
        elseif ($code -match '^\s*MACROBUTTON\s+MTEditEquationSection') { $breaks += $row }
        Release-ComObject $field
    }
    return [ordered]@{ numbers = $numbers; references = $refs; breaks = $breaks;
        mathtype_objects = @($script:Document.InlineShapes | Where-Object { $_.Type -eq 1 -and $_.OLEFormat.ProgID -eq 'Equation.DSMT4' }).Count;
        omath_objects = $script:Document.OMaths.Count;
        word_sections = $script:Document.Sections.Count;
        format = [string]$script:Word.Run('MTEqnNumFormat.GetEqnNumFormat', $script:Document) }
}

function Assert-NativeSnapshot($snapshot, [string[]]$ExpectedNumbers, [string[]]$ExpectedReferences, [int]$ExpectedBreaks) {
    if (($snapshot.numbers.text -join ',') -ne ($ExpectedNumbers -join ',')) { throw "Numbers differ: $($snapshot.numbers.text -join ',')" }
    if (($snapshot.references.text -join ',') -ne ($ExpectedReferences -join ',')) { throw "References differ: $($snapshot.references.text -join ',')" }
    if ($snapshot.mathtype_objects -ne 6 -or $snapshot.omath_objects -ne 0) { throw 'Expected six genuine MathType objects and no OMath.' }
    if ($snapshot.breaks.Count -ne $ExpectedBreaks -or $snapshot.word_sections -ne 1) { throw 'MathType equation breaks or Word-section count differs.' }
    foreach ($row in $snapshot.numbers) { if ($row.code -notmatch 'SEQ\s+MTEqn' -or $row.code -notmatch 'MTPlaceRef') { throw 'Missing native number structure.' } }
    foreach ($row in $snapshot.references) { if (-not $row.target_exists -or $row.code -notmatch '\bREF\s+ZEqnNum') { throw 'Missing native reference structure.' } }
}

function Save-NativeSample([string]$Path, $ExpectedNumbers, $ExpectedReferences, [int]$ExpectedBreaks) {
    $script:Selection.SetRange(0, 0)
    $null = $script:Word.Run('MTCommand_UpdateEqns')
    $null = $script:Document.Fields.Update()
    $script:Document.SaveAs2($Path, 16)
    $script:Document.Close(0)
    $script:Document = $script:Word.Documents.Open($Path, $false, $true, $false)
    $script:Selection = $script:Word.Selection
    $snapshot = Get-NativeSnapshot
    Assert-NativeSnapshot $snapshot $ExpectedNumbers $ExpectedReferences $ExpectedBreaks
    $script:Document.PrintRevisions = $false
    $script:Document.ExportAsFixedFormat([IO.Path]::ChangeExtension($Path, '.pdf'), 17)
    return $snapshot
}

$samples = @(
    @{ name='连续编号'; format='(#E1)'; section=$false; chapter=$false; numbers=@('(1)','(2)','(3)','(4)','(5)','(6)'); refs=@('(1)','(2)','(3)','(4)','(6)','(1)'); breaks=1 },
    @{ name='按章编号'; format='(#C1-#E1)'; section=$false; chapter=$true; numbers=@('(1-1)','(1-2)','(1-3)','(2-1)','(2-2)','(2-3)'); refs=@('(1-1)','(1-2)','(1-3)','(2-1)','(2-3)','(1-1)'); breaks=2 },
    @{ name='按章和节编号'; format='(#C1.#S1.#E1)'; section=$true; chapter=$true; numbers=@('(1.1.1)','(1.1.2)','(1.2.1)','(2.1.1)','(2.1.2)','(2.2.1)'); refs=@('(1.1.1)','(1.1.2)','(1.2.1)','(2.1.1)','(2.2.1)','(1.1.1)'); breaks=4 }
)
$tex = @('x+y=z', 'x=\frac{-b+\sqrt{b^2-4ac}}{2a}', 'E=mc^2', 'F=ma', 's=vt', 'P=\frac{W}{t}')
$folder = [IO.Path]::GetFullPath($OutputPath)
if ((Test-Path -LiteralPath $folder) -and @(Get-ChildItem -LiteralPath $folder -Force).Count) { throw 'Use a new or empty output folder for this test.' }
$null = New-Item -ItemType Directory -Path $folder -Force
$reports = @()
try {
    Enter-BridgeLock
    Set-TemporaryMathTypePreferences
    foreach ($sample in $samples) {
        Start-Word
        $script:Document.PageSetup.PaperSize = 7 # A4
        $script:Document.PageSetup.TopMargin = 50
        $script:Document.PageSetup.BottomMargin = 50
        $normal = $script:Document.Styles.Item(-1)
        $normal.Font.Name = 'Times New Roman'; $normal.Font.NameFarEast = '宋体'; $normal.Font.Size = 10.5
        $normal.ParagraphFormat.SpaceAfter = 5
        $normal.ParagraphFormat.LineSpacingRule = 0
        $body = @('MathType 原生公式编号测试', "测试方式 $($sample.name)",
            '以下公式编号及正文引用由本机 MathType 命令生成。可双击公式编辑内容，双击正文引用定位目标编号。',
            '第一章 第一节', '{{MATH:e1}}', '由式 {{EQREF:r1}} 可知，x 与 y 的和为 z。',
            '{{MATH:e2}}', '式 {{EQREF:r2}} 展示分式和根号。', '{{BREAK:s1}}', '第一章 第二节',
            '{{MATH:e3}}', '式 {{EQREF:r3}} 使用当前章节的编号。', '{{BREAK:c2}}', '第二章 第一节',
            '{{MATH:e4}}', '式 {{EQREF:r4}} 用于检查换章后的编号。', '{{MATH:e5}}',
            '{{BREAK:s2}}', '第二章 第二节', '{{MATH:e6}}', '式 {{EQREF:r6}} 用于检查换节后的编号。',
            '跨章引用仍指向第一条公式，见式 {{EQREF:cross}}。')
        $script:Document.Content.Text = ($body -join "`r") + "`r"
        $title = $script:Document.Paragraphs.Item(1).Range
        $title.Style = $script:Document.Styles.Item(-63) # Title
        $title.Font.NameFarEast = '黑体'; $title.Font.Size = 16; $title.Font.Color = 0
        # MathType reads this document-level preference itself. No global default is changed.
        $code = $script:Word.Run('MTLib.WriteDocPropString', $script:Document, 'MTEquationNumber2', $sample.format)
        if ($code -ne 0) { throw 'MathType could not save the document-local number format.' }
        $fields = @{}
        foreach ($event in @('e1','e2','s1','e3','c2','e4','e5','s2','e6')) {
            if ($event.StartsWith('e')) {
                $index = [int]$event.Substring(1) - 1
                $equation = [pscustomobject]@{ id=$event; marker="{{MATH:$event}}"; tex=$tex[$index]; layout='display'; numbered=$true }
                $number = Convert-Equation -Equation $equation -RepairExactHeights
                $fields[$event] = $number
                $script:HeldComObjects.Add($number)
                $para = $number.Result.Paragraphs.Item(1).Range
                $null = $script:Document.Bookmarks.Add("Demo_$event", $para)
            }
            else {
                $range = Find-UniqueMarkerRange -Marker "{{BREAK:$event}}"
                $pos = $range.Start; $range.Text = ''; $script:Selection.SetRange($pos,$pos)
                if ($event -eq 'c2' -and $sample.chapter) { $null = $script:Word.Run('MTCommand_InsertNextChapter') }
                elseif ($event -ne 'c2' -and $sample.section) { $null = $script:Word.Run('MTCommand_InsertNextSection') }
                Release-ComObject $range
            }
        }
        foreach ($id in @('e1','e2','e3','e4','e6')) {
            $ref = [pscustomobject]@{ marker='{{EQREF:r' + $id.Substring(1) + '}}'; target=$id }
            Insert-NativeReference -Reference $ref -TargetField $fields[$id]
        }
        Insert-NativeReference -Reference ([pscustomobject]@{marker='{{EQREF:cross}}';target='e1'}) -TargetField $fields.e1
        $path = Join-Path $folder ($sample.name + '-mathtype.docx')
        $report = Save-NativeSample $path $sample.numbers $sample.refs $sample.breaks
        $reports += @{ name=$sample.name; path=$path; saved=$report }
        $script:Document.Close(0); $script:Document = $null
        if ($sample.section) {
            $script:Document = $script:Word.Documents.Open($path, $false, $false, $false)
            $script:Selection = $script:Word.Selection
            # Cut/paste the second formula paragraph before the first. Native reference bookmarks move with it.
            $move = $script:Document.Bookmarks.Item('Demo_e2').Range
            $move.Cut()
            $destination = $script:Document.Bookmarks.Item('Demo_e1').Range
            $destination.Collapse(1); $destination.Paste()
            $movedPath = Join-Path $folder '调整公式顺序后-mathtype.docx'
            $movedRefs = @('(1.1.2)','(1.1.1)','(1.2.1)','(2.1.1)','(2.2.1)','(1.1.2)')
            $moved = Save-NativeSample $movedPath $sample.numbers $movedRefs $sample.breaks
            $reports += @{name='调整公式顺序后';path=$movedPath;saved=$moved}
        }
        Stop-Word
    }
    $result = [ordered]@{ok=$true; samples=$reports; method='Installed MathType macros; document-local MTEquationNumber2; no simple-number normalization'}
    $result | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $folder 'mathtype-for-word-numbering-report.json') -Encoding utf8
    Write-JsonResult $result
}
catch { Write-JsonResult @{ok=$false;error=$_.Exception.ToString();location=$_.InvocationInfo.PositionMessage;stack=$_.ScriptStackTrace}; exit 1 }
finally { Stop-Word; Restore-MathTypePreferences; Exit-BridgeLock }
