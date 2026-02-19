param(
  [string]$InputWorkbook = ".\macros_test.xls",
  [string]$PatchModule = ".\deliverables\Cisco_DeliveryMode_Patch.bas",
  [string]$OutputWorkbook = ".\macros_test_ready.xlsm"
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  if (-not (Test-Path $InputWorkbook)) { throw "Input workbook not found: $InputWorkbook" }
  if (-not (Test-Path $PatchModule)) { throw "Patch module not found: $PatchModule" }

  $wb = $excel.Workbooks.Open((Resolve-Path $InputWorkbook).Path)

  # Remove duplicate GO_CISCO if present in known modules
  foreach ($component in @($wb.VBProject.VBComponents)) {
    try {
      $code = $component.CodeModule
      if ($code -and $code.CountOfLines -gt 0) {
        $all = $code.Lines(1, $code.CountOfLines)
        if ($all -match "Public Sub GO_CISCO\(\)" -or $all -match "Sub GO_CISCO\(\)") {
          # Keep old code untouched; user can decide. We only import patched module.
        }
      }
    } catch {}
  }

  # Import patch module
  $wb.VBProject.VBComponents.Import((Resolve-Path $PatchModule).Path) | Out-Null

  # Ensure MACROS sheet has dropdown source hints in helper cells (safe fallback)
  $ws = $wb.Worksheets.Item("MACROS")
  $ws.Range("Z1").Value2 = "CIP"
  $ws.Range("Z2").Value2 = "DDP"
  $ws.Range("Z3").Value2 = "CIP&DDP"
  $ws.Range("B10").Value2 = "CIP"

  # Save as macro-enabled workbook
  $xlOpenXMLWorkbookMacroEnabled = 52
  $out = (Resolve-Path ".").Path + "\\" + (Split-Path $OutputWorkbook -Leaf)
  $wb.SaveAs($out, $xlOpenXMLWorkbookMacroEnabled)
  $wb.Close($true)

  Write-Host "READY: $out"
}
finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
