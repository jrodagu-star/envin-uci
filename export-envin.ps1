# Extrae ENVIN-PAZ a envin-data.js (UTF-8) para el explorador local.
$ErrorActionPreference = "Stop"
$src = "C:\Users\jroda\Downloads\ENVINPAZ.accdb"
$out = Join-Path $PSScriptRoot "envin-data.js"

function Esc([string]$s) {
  if ($null -eq $s) { return "" }
  $s = $s.Replace("\", "\\").Replace('"', '\"').Replace("`r", " ").Replace("`n", " ").Replace("`t", " ")
  return $s
}
function IsBlank($v) {
  return ($null -eq $v) -or ($v -is [DBNull]) -or ([string]$v).Trim() -eq ""
}
function AsDate($v) {
  if (IsBlank $v) { return $null }
  try { return ([datetime]$v).ToString("yyyy-MM-dd") } catch { return $null }
}
function AsInt($v) {
  if (IsBlank $v) { return $null }
  $n = 0
  if ([int]::TryParse(([string]$v).Trim(), [ref]$n)) { return $n }
  return $null
}
function AsNum($v) {
  if (IsBlank $v) { return $null }
  $n = 0.0
  if ([double]::TryParse(([string]$v).Trim(), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
  $n2 = 0.0
  if ([double]::TryParse(([string]$v).Trim(), [ref]$n2)) { return $n2 }
  return $null
}
function Yes1($v) {
  # ENVIN: 1 = Si, 2 = No. Boolean True = Si.
  if (IsBlank $v) { return $null }
  if ($v -is [bool]) { return $(if ($v) { 1 } else { 0 }) }
  $s = ([string]$v).Trim().ToLowerInvariant()
  if ($s -eq "1" -or $s -eq "true" -or $s -eq "-1") { return 1 }
  if ($s -eq "2" -or $s -eq "0" -or $s -eq "false") { return 0 }
  return $null
}
function Jnum($v) { if ($null -eq $v) { "null" } else { [string]$v } }
function Jstr($v) { if ($null -eq $v) { "null" } else { '"' + (Esc ([string]$v)) + '"' } }

$conn = New-Object -ComObject ADODB.Connection
$conn.Open("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$src;Persist Security Info=False;")

function Get-Map($sql, $codeCol, $nameCol) {
  $map = [ordered]@{}
  $rs = $conn.Execute($sql)
  while (-not $rs.EOF) {
    $c = [string]$rs.Fields.Item($codeCol).Value
    $n = [string]$rs.Fields.Item($nameCol).Value
    if (-not [string]::IsNullOrWhiteSpace($c)) { $map[$c.Trim()] = $n }
    $rs.MoveNext()
  }
  $rs.Close()
  return $map
}

Write-Host "Catalogos..."
$infec = Get-Map "SELECT CODIGO, NOMBRE FROM [dbo_COD_INF]" "CODIGO" "NOMBRE"
$dx = Get-Map "SELECT CODIGO, NOMBRE FROM [dbo_COD_DX]" "CODIGO" "NOMBRE"
$atbCat = Get-Map "SELECT CODIGO, NOMBRE FROM [dbo_COD_ATB]" "CODIGO" "NOMBRE"
# dbo_COD_ORIGEN second column has special char; read by index
$origenPac = [ordered]@{}
$rs = $conn.Execute("SELECT * FROM [dbo_COD_ORIGEN]")
while (-not $rs.EOF) {
  $origenPac[[string]$rs.Fields.Item(0).Value] = [string]$rs.Fields.Item(1).Value
  $rs.MoveNext()
}
$rs.Close()

$tipoIng = [ordered]@{}
$rs = $conn.Execute("SELECT * FROM [dbo_COD_TIPO_INGRESO]")
while (-not $rs.EOF) {
  $tipoIng[[string]$rs.Fields.Item(0).Value] = [string]$rs.Fields.Item(1).Value
  $rs.MoveNext()
}
$rs.Close()

$grupos = [ordered]@{}
$rs = $conn.Execute("SELECT * FROM [dbo_GRUP_GERMEN]")
while (-not $rs.EOF) {
  $grupos[[string]$rs.Fields.Item(0).Value] = [string]$rs.Fields.Item(1).Value
  $rs.MoveNext()
}
$rs.Close()

$indicAtb = [ordered]@{}
$rs = $conn.Execute("SELECT * FROM [dbo_COD_INDICACION_ATB]")
while (-not $rs.EOF) {
  $indicAtb[[string]$rs.Fields.Item(0).Value] = [string]$rs.Fields.Item(1).Value
  $rs.MoveNext()
}
$rs.Close()

$mues = [ordered]@{}
$rs = $conn.Execute("SELECT * FROM [dbo_COD_MUES]")
while (-not $rs.EOF) {
  $cod = [string]$rs.Fields.Item("CODIGO").Value
  $nom = [string]$rs.Fields.Item("NOMBRE").Value
  $grp = ([string]$rs.Fields.Item("GrupoInfec").Value).Trim()
  if ($grp -eq "V" -or -not $mues.Contains($cod)) { $mues[$cod] = $nom }
  $rs.MoveNext()
}
$rs.Close()

$ger = [ordered]@{}
$rs = $conn.Execute("SELECT CODIGO, Nombre, Grupo, Gram FROM [dbo_COD_GER]")
while (-not $rs.EOF) {
  $cod = [string]$rs.Fields.Item("CODIGO").Value
  $ger[$cod] = @{
    n = [string]$rs.Fields.Item("Nombre").Value
    g = [string]$rs.Fields.Item("Grupo").Value
    gram = [string]$rs.Fields.Item("Gram").Value
  }
  $rs.MoveNext()
}
$rs.Close()

Write-Host "Ingresos..."
$ingresos = New-Object System.Collections.Generic.List[string]
$rs = $conn.Execute("SELECT NHC, INICIALES, EDAD, SEXO, ING_UCI, ALTA_UCI, EXITUS, DIAG, APACHE, SAPSII, ORIGEN, ADMISION, VM, CVC, SONDAURI, INMUNOSU, NEUTROPE, INMUNODE, TRAUMA, CORONARIO, ATB48H, pacienteCovid19, INSUFREN, EPOC, CIRROSIS, NEOPLASIA, DIABETES, MARSA, ACINETO, BLEES, PSEUDOMONAS, ERV, BGNMR, Peso, GLASGOW_Est, idUnico, TIPOENVIN FROM [dbo_INICIAL]")
$nIng = 0
while (-not $rs.EOF) {
  $nhc = [string]$rs.Fields.Item("NHC").Value
  $ing = AsDate $rs.Fields.Item("ING_UCI").Value
  $alta = AsDate $rs.Fields.Item("ALTA_UCI").Value
  $anio = $null
  if ($ing) { $anio = [int]$ing.Substring(0,4) }
  $ap = AsNum $rs.Fields.Item("APACHE").Value
  $sa = AsNum $rs.Fields.Item("SAPSII").Value
  $ex = AsInt $rs.Fields.Item("EXITUS").Value
  $sx = AsInt $rs.Fields.Item("SEXO").Value
  $obj = '{' +
    '"nhc":' + (Jstr $nhc) + ',' +
    '"ini":' + (Jstr $rs.Fields.Item("INICIALES").Value) + ',' +
    '"ed":' + (Jnum (AsInt $rs.Fields.Item("EDAD").Value)) + ',' +
    '"sx":' + (Jnum $sx) + ',' +
    '"f0":' + (Jstr $ing) + ',' +
    '"f1":' + (Jstr $alta) + ',' +
    '"y":' + (Jnum $anio) + ',' +
    '"ex":' + (Jnum $ex) + ',' +
    '"dx":' + (Jnum (AsInt $rs.Fields.Item("DIAG").Value)) + ',' +
    '"ap":' + (Jnum $ap) + ',' +
    '"sa":' + (Jnum $sa) + ',' +
    '"or":' + (Jnum (AsInt $rs.Fields.Item("ORIGEN").Value)) + ',' +
    '"ad":' + (Jnum (AsInt $rs.Fields.Item("ADMISION").Value)) + ',' +
    '"vm":' + (Jnum (Yes1 $rs.Fields.Item("VM").Value)) + ',' +
    '"cvc":' + (Jnum (Yes1 $rs.Fields.Item("CVC").Value)) + ',' +
    '"su":' + (Jnum (Yes1 $rs.Fields.Item("SONDAURI").Value)) + ',' +
    '"im":' + (Jnum (Yes1 $rs.Fields.Item("INMUNOSU").Value)) + ',' +
    '"tr":' + (Jnum (Yes1 $rs.Fields.Item("TRAUMA").Value)) + ',' +
    '"co":' + (Jnum (Yes1 $rs.Fields.Item("CORONARIO").Value)) + ',' +
    '"cov":' + (Jnum (Yes1 $rs.Fields.Item("pacienteCovid19").Value)) + ',' +
    '"ir":' + (Jnum (Yes1 $rs.Fields.Item("INSUFREN").Value)) + ',' +
    '"epoc":' + (Jnum (Yes1 $rs.Fields.Item("EPOC").Value)) + ',' +
    '"cir":' + (Jnum (Yes1 $rs.Fields.Item("CIRROSIS").Value)) + ',' +
    '"neo":' + (Jnum (Yes1 $rs.Fields.Item("NEOPLASIA").Value)) + ',' +
    '"dm":' + (Jnum (Yes1 $rs.Fields.Item("DIABETES").Value)) + ',' +
    '"marsa":' + (Jnum (Yes1 $rs.Fields.Item("MARSA").Value)) + ',' +
    '"aci":' + (Jnum (Yes1 $rs.Fields.Item("ACINETO").Value)) + ',' +
    '"blee":' + (Jnum (Yes1 $rs.Fields.Item("BLEES").Value)) + ',' +
    '"pse":' + (Jnum (Yes1 $rs.Fields.Item("PSEUDOMONAS").Value)) + ',' +
    '"peso":' + (Jnum (AsNum $rs.Fields.Item("Peso").Value)) + ',' +
    '"gcs":' + (Jnum (AsInt $rs.Fields.Item("GLASGOW_Est").Value)) +
    '}'
  $ingresos.Add($obj) | Out-Null
  $nIng++
  $rs.MoveNext()
}
$rs.Close()
Write-Host "  $nIng ingresos"

Write-Host "Microorganismos..."
$microByInf = @{}
$rs = $conn.Execute("SELECT NHC, ING_UCI, F_INF, INFEC, MICRO, UFC, IMPORTADO FROM [dbo_MICROORGANISMOS]")
$nMic = 0
while (-not $rs.EOF) {
  $nhc = ([string]$rs.Fields.Item("NHC").Value).Trim()
  $ing = AsDate $rs.Fields.Item("ING_UCI").Value
  $fi = AsDate $rs.Fields.Item("F_INF").Value
  $inf = AsInt $rs.Fields.Item("INFEC").Value
  $key = "$nhc|$ing|$fi|$inf"
  $m = AsInt $rs.Fields.Item("MICRO").Value
  $u = AsInt $rs.Fields.Item("UFC").Value
  $imp = AsInt $rs.Fields.Item("IMPORTADO").Value
  $item = '{"m":' + (Jnum $m) + ',"u":' + (Jnum $u) + ',"imp":' + (Jnum $imp) + '}'
  if (-not $microByInf.ContainsKey($key)) { $microByInf[$key] = New-Object System.Collections.Generic.List[string] }
  $microByInf[$key].Add($item) | Out-Null
  $nMic++
  $rs.MoveNext()
}
$rs.Close()
Write-Host "  $nMic aislamientos"

Write-Host "Infecciones..."
$infecciones = New-Object System.Collections.Generic.List[string]
$rs = $conn.Execute("SELECT NHC, ING_UCI, F_INF, INFEC, ORIG, MUES, DX_CLIN, TTOATB, Bacteriemia, TTOATBApropiado FROM [dbo_INFECCIO]")
$nInf = 0
while (-not $rs.EOF) {
  $nhc = ([string]$rs.Fields.Item("NHC").Value).Trim()
  $ing = AsDate $rs.Fields.Item("ING_UCI").Value
  $fi = AsDate $rs.Fields.Item("F_INF").Value
  $inf = AsInt $rs.Fields.Item("INFEC").Value
  $anio = $null
  if ($fi) { $anio = [int]$fi.Substring(0,4) } elseif ($ing) { $anio = [int]$ing.Substring(0,4) }
  $key = "$nhc|$ing|$fi|$inf"
  $ms = "[]"
  if ($microByInf.ContainsKey($key)) { $ms = "[" + ($microByInf[$key] -join ",") + "]" }
  $bac = $rs.Fields.Item("Bacteriemia").Value
  $obj = '{' +
    '"nhc":' + (Jstr $nhc) + ',' +
    '"f0":' + (Jstr $ing) + ',' +
    '"fi":' + (Jstr $fi) + ',' +
    '"y":' + (Jnum $anio) + ',' +
    '"inf":' + (Jnum $inf) + ',' +
    '"orig":' + (Jnum (AsInt $rs.Fields.Item("ORIG").Value)) + ',' +
    '"mues":' + (Jnum (AsInt $rs.Fields.Item("MUES").Value)) + ',' +
    '"dxc":' + (Jnum (AsInt $rs.Fields.Item("DX_CLIN").Value)) + ',' +
    '"atb":' + (Jnum (AsInt $rs.Fields.Item("TTOATB").Value)) + ',' +
    '"bac":' + (Jnum (Yes1 $bac)) + ',' +
    '"ms":' + $ms +
    '}'
  $infecciones.Add($obj) | Out-Null
  $nInf++
  $rs.MoveNext()
}
$rs.Close()
Write-Host "  $nInf infecciones"

Write-Host "Antibioticos..."
$atbs = New-Object System.Collections.Generic.List[string]
$rs = $conn.Execute("SELECT NHC, ING_UCI, ATB, I_ATB, F_ATB, INDIC1, InfeccionAsociada FROM [dbo_ATB]")
$nAtb = 0
while (-not $rs.EOF) {
  $obj = '{' +
    '"nhc":' + (Jstr ([string]$rs.Fields.Item("NHC").Value).Trim()) + ',' +
    '"f0":' + (Jstr (AsDate $rs.Fields.Item("ING_UCI").Value)) + ',' +
    '"atb":' + (Jnum (AsInt $rs.Fields.Item("ATB").Value)) + ',' +
    '"i":' + (Jstr (AsDate $rs.Fields.Item("I_ATB").Value)) + ',' +
    '"f":' + (Jstr (AsDate $rs.Fields.Item("F_ATB").Value)) + ',' +
    '"ind":' + (Jnum (AsInt $rs.Fields.Item("INDIC1").Value)) +
    '}'
  $atbs.Add($obj) | Out-Null
  $nAtb++
  $rs.MoveNext()
}
$rs.Close()
Write-Host "  $nAtb tratamientos"
$conn.Close()

function MapJson($map) {
  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($k in $map.Keys) {
    $parts.Add(('"' + (Esc $k) + '":"' + (Esc $map[$k]) + '"')) | Out-Null
  }
  return "{" + ($parts -join ",") + "}"
}
function GerJson($map) {
  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($k in $map.Keys) {
    $g = $map[$k]
    $parts.Add(('"' + (Esc $k) + '":{"n":"' + (Esc $g.n) + '","g":"' + (Esc $g.g) + '","gram":"' + (Esc $g.gram) + '"}')) | Out-Null
  }
  return "{" + ($parts -join ",") + "}"
}

$origInf = @{
  "1" = "Comunitaria"
  "2" = "Nosocomial UCI"
  "3" = "Nosocomial extra-UCI"
  "4" = "Relacionada con cuidados sanitarios"
}
$sexo = @{ "1" = "Hombre"; "2" = "Mujer" }
$exitus = @{ "1" = "Fallecido en UCI"; "2" = "Vivo al alta de UCI" }
$gram = @{ "N" = "Gram negativo"; "P" = "Gram positivo"; "H" = "Hongo"; "V" = "Virus"; "C" = "Cultivo negativo"; "O" = "Otro" }

Write-Host "Escribiendo $out ..."
$sw = New-Object System.IO.StreamWriter($out, $false, (New-Object System.Text.UTF8Encoding $false))
$sw.Write("window.ENVIN=")
$sw.Write("{")
$sw.Write('"meta":{"fuente":"ENVINPAZ.accdb","extraido":"' + (Get-Date -Format "yyyy-MM-dd") + '","ingresos":' + $nIng + ',"infecciones":' + $nInf + ',"aislamientos":' + $nMic + ',"atb":' + $nAtb + ',"periodo":[2009,2026]},')
$sw.Write('"cat":{')
$sw.Write('"infec":' + (MapJson $infec) + ',')
$sw.Write('"dx":' + (MapJson $dx) + ',')
$sw.Write('"atb":' + (MapJson $atbCat) + ',')
$sw.Write('"origenPac":' + (MapJson $origenPac) + ',')
$sw.Write('"tipoIng":' + (MapJson $tipoIng) + ',')
$sw.Write('"grupos":' + (MapJson $grupos) + ',')
$sw.Write('"indicAtb":' + (MapJson $indicAtb) + ',')
$sw.Write('"mues":' + (MapJson $mues) + ',')
$sw.Write('"origInf":' + (MapJson $origInf) + ',')
$sw.Write('"sexo":' + (MapJson $sexo) + ',')
$sw.Write('"exitus":' + (MapJson $exitus) + ',')
$sw.Write('"gram":' + (MapJson $gram) + ',')
$sw.Write('"micro":' + (GerJson $ger))
$sw.Write("},")
$sw.Write('"ingresos":[')
$sw.Write(($ingresos -join ","))
$sw.Write("],")
$sw.Write('"infecciones":[')
$sw.Write(($infecciones -join ","))
$sw.Write("],")
$sw.Write('"atb":[')
$sw.Write(($atbs -join ","))
$sw.Write("]")
$sw.Write("};")
$sw.Close()
Write-Host ("OK " + ((Get-Item $out).Length / 1MB).ToString("0.00") + " MB")
