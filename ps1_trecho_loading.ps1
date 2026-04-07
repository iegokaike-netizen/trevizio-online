# TRECHO PARA PROGRESSO

$total = $clientes.Count
$contador = 0

foreach ($cliente in $clientes) {

    $contador++

    $percent = [math]::Round(($contador / $total) * 100)

    $resultado = @{
        progresso = $percent
        cliente = $cliente
    }

    $resultado | ConvertTo-Json -Depth 5
}
