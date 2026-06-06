# actualizar.ps1 - Lee GOOGLE DASH.xlsx y genera data.js para el dashboard

$ExcelFile = 'C:\Users\CALL JEFE\OneDrive - equilibriototal.com.mx 1\GOOGLE DASH.xlsx'
$OutputFile = Join-Path $PSScriptRoot 'data.js'

if (-not (Test-Path $ExcelFile)) {
    Write-Host "ERROR: No se encontro el archivo: $ExcelFile"
    exit 1
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($ExcelFile, 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $lastRow = $ws.UsedRange.Rows.Count

    $C = @{ dia=1; nombre=2; numero=3; agenda=4; sede=5; asiste=6; costo=7; plan=8; monto=9; cxc=10 }

    $rows = [System.Collections.Generic.List[object]]::new()

    for ($r = 2; $r -le $lastRow; $r++) {
        $nombre = $ws.Cells.Item($r, $C.nombre).Text.Trim()
        if (-not $nombre) { continue }

        $sedeRaw   = $ws.Cells.Item($r, $C.sede).Text.Trim()
        $asisteRaw = $ws.Cells.Item($r, $C.asiste).Text.Trim().ToUpper()
        $planRaw   = $ws.Cells.Item($r, $C.plan).Text.Trim().ToUpper()
        $diaRaw    = $ws.Cells.Item($r, $C.dia).Text.Trim()
        $numeroRaw = $ws.Cells.Item($r, $C.numero).Text.Trim()
        $agendaRaw = $ws.Cells.Item($r, $C.agenda).Text.Trim()
        $costoTxt  = $ws.Cells.Item($r, $C.costo).Text.Trim()

        $costoVal = $ws.Cells.Item($r, $C.costo).Value2
        $montoVal = $ws.Cells.Item($r, $C.monto).Value2
        $cxcVal   = $ws.Cells.Item($r, $C.cxc).Value2

        $costo = if ($costoVal -is [double] -or $costoVal -is [int]) { [double]$costoVal } else { 0.0 }
        $monto = if ($montoVal -is [double] -or $montoVal -is [int]) { [double]$montoVal } else { 0.0 }
        $cxc   = if ($cxcVal  -is [double] -or $cxcVal  -is [int]) { [double]$cxcVal  } else { 0.0 }

        if ($costoTxt -match 'NO.*PAGO|SIN.*PAGO') { $costo = 0.0 }

        $hasPlan = $planRaw -match '^S'

        $montosLiquidados = @(4900.0, 5900.0, 6900.0)

        if (($cxc -eq 0) -and $hasPlan -and ($monto -gt 0) -and ($monto -notin $montosLiquidados)) {
            $cxc = 5900.0 - $monto
        }

        if ($asisteRaw -match 'PEND') { $asiste = 'PENDIENTE' }
        elseif ($asisteRaw -match '^S') { $asiste = 'SI' }
        else { $asiste = 'NO' }

        $irrs = [System.Collections.Generic.List[string]]::new()
        if ($costoTxt -match '\?' -or ($costo -eq 0 -and $costoTxt -ne '' -and $costoTxt -notmatch 'PAGO')) {
            $irrs.Add("Costo invalido: $costoTxt")
        } elseif ($asiste -eq 'SI' -and $costo -gt 0 -and $costo -notin @(490.0, 590.0, 690.0)) {
            $costoInt = [int]$costo
            $irrs.Add("Consulta `$$costoInt (esperado 590/690)")
        }
        if ($hasPlan -and $monto -gt 0 -and ($monto -notin $montosLiquidados)) {
            $montoInt = [int]$monto
            $irrs.Add("Plan `$$montoInt (esperado 4900/5900/6900)")
        }

        $nombreClean = $nombre -replace '"', '' -replace "'", ''

        $rows.Add([PSCustomObject]@{
            dia     = $diaRaw
            agenda  = $agendaRaw
            nombre  = $nombreClean
            numero  = $numeroRaw
            sede    = $sedeRaw
            asiste  = $asiste
            costo   = $costo
            hasPlan = $hasPlan
            monto   = $monto
            cxc     = $cxc
            irrs    = @($irrs)
        })
    }

    # Leer PP (inversión) desde hoja "pp" — B1=Mayo, B2=Junio
    $ppMayo = 0; $ppJunio = 0
    try {
        $wsPP = $null
        foreach ($sheet in $wb.Worksheets) {
            if ($sheet.Name.ToUpper() -eq 'PP') { $wsPP = $sheet; break }
        }
        if ($wsPP) {
            $v1 = $wsPP.Cells.Item(1, 2).Value2
            $v2 = $wsPP.Cells.Item(2, 2).Value2
            if ($v1 -is [double] -and $v1 -gt 0) { $ppMayo  = [double]$v1 }
            if ($v2 -is [double] -and $v2 -gt 0) { $ppJunio = [double]$v2 }
        }
    } catch {}
    $ppVal = $ppMayo + $ppJunio

    # Leer Leads desde hoja "cc" — B1=Mayo, B2=Junio
    $leadsMayo = 0; $leadsJunio = 0
    try {
        $wsCC = $null
        foreach ($sheet in $wb.Worksheets) {
            if ($sheet.Name.ToUpper() -eq 'CC') { $wsCC = $sheet; break }
        }
        if ($wsCC) {
            $v1 = $wsCC.Cells.Item(1, 2).Value2
            $v2 = $wsCC.Cells.Item(2, 2).Value2
            if ($v1 -is [double] -and $v1 -gt 0) { $leadsMayo  = [int]$v1 }
            if ($v2 -is [double] -and $v2 -gt 0) { $leadsJunio = [int]$v2 }
        }
    } catch {}
    $leadsVal = $leadsMayo + $leadsJunio

    $wb.Close($false)

    $json = (@($rows) | ConvertTo-Json -Depth 4 -Compress)
    $ts   = Get-Date -Format 'dd/MM/yyyy HH:mm'
    $cnt  = $rows.Count

    $extraLines = ""
    if ($ppVal    -gt 0) { $extraLines += "`nwindow.DASHBOARD_PP = $ppVal;" }
    if ($ppMayo   -gt 0) { $extraLines += "`nwindow.DASHBOARD_PP_MAYO = $ppMayo;" }
    if ($ppJunio  -gt 0) { $extraLines += "`nwindow.DASHBOARD_PP_JUNIO = $ppJunio;" }
    if ($leadsVal  -gt 0) { $extraLines += "`nwindow.DASHBOARD_LEADS = $leadsVal;" }
    if ($leadsMayo -gt 0) { $extraLines += "`nwindow.DASHBOARD_LEADS_MAYO = $leadsMayo;" }
    if ($leadsJunio -gt 0){ $extraLines += "`nwindow.DASHBOARD_LEADS_JUNIO = $leadsJunio;" }

    $content = "// Auto-generado $ts - $cnt pacientes`nwindow.DASHBOARD_DATA = $json;`nwindow.DASHBOARD_FILE = 'GOOGLE DASH.xlsx';`nwindow.DASHBOARD_TS = '$ts';`nwindow.DASHBOARD_CNT = $cnt;$extraLines"

    [System.IO.File]::WriteAllText($OutputFile, $content, [System.Text.Encoding]::UTF8)
    Write-Host "OK: $cnt pacientes - PP Mayo: $ppMayo / Jun: $ppJunio - Leads Mayo: $leadsMayo / Jun: $leadsJunio - $ts"

} finally {
    if ($excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
}
