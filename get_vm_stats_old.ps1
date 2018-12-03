Connect-VIServer vcsa1-vflab.oss.vf.rogers.com

$vm= Get-Content /export/home/john.rumley/server
$file = "/export/home/john.rumley/report.xlsx"

Remove-Item $file -ErrorAction Ignore

$entities = $vm
$start = (Get-Date).AddDays(-90)
$finish = (Get-Date).AddDays(-1)
$interval = 1
$entities = $vm
$stat = 'cpu.usagemhz.average','cpu.usage.average','mem.consumed.average','mem.usage.average','net.usage.average'

$result =
ForEach($item in $vm)
{Get-VM -Name $item | Select Name,NumCpu,MemoryMB}
$result | Export-Excel -Path "$file" -WorksheetName "VM Specifications" -AutoNameRange


Get-Stat -Entity $entities -Stat $stat -Start $start -Finish $finish |
Group-Object -Property Entity | %{
    $report = $_.Group | Group-Object -Property Timestamp | %{
        $obj = [ordered]@{
            Timestamp = $_.Name
            CPU_Usage_MHz = $_.Group | where{$_.MetricId -eq 'cpu.usagemhz.average'} | select -ExpandProperty Value
            CPU_Usage_Percent = $_.Group | where{$_.MetricId -eq 'cpu.usage.average'} | select -ExpandProperty Value
            Memory_Usage_KB = $_.Group | where{$_.MetricId -eq 'mem.consumed.average'} | select -ExpandProperty Value
            Memory_Usage_Percent = $_.Group | where{$_.MetricId -eq 'mem.usage.average'} | select -ExpandProperty Value
            Network_Usage_KBps = $_.Group | where{$_.MetricId -eq 'net.usage.average'} | select -ExpandProperty Value
        }
        New-Object PSObject -Property $obj
    }

    $chart_cpu = New-ExcelChartDefinition -ChartType line -XRange "Timestamp" -YRange "CPU_Usage_MHz","CPU_Usage_Percent" -Title "Average CPU Usage" -Width 800 -TitleBold -TitleSize 14 -XAxisTitleText "Date" -XAxisTitleBold -XAxisTitleSize 12 -YAxisTitleBold -YAxisTitleSize 12 -SeriesHeader "MHz","%" -Row 0 -Column 10
    $chart_mem = New-ExcelChartDefinition -ChartType line -XRange "Timestamp" -YRange "Memory_Usage_KB","Memory_Usage_Percent" -Title "Average Memory Consumed - KiloBytes" -Width 800 -TitleBold -TitleSize 14 -XAxisTitleText "Date" -XAxisTitleBold -XAxisTitleSize 12 -YAxisTitleBold -YAxisTitleSize 12 -SeriesHeader "KB","%" -Row 20 -Column 10
    $chart_net = New-ExcelChartDefinition -ChartType line -XRange "Timestamp" -YRange "Memory_Usage_KB" -Title "Average Network Usage - KBps" -Width 800 -TitleBold -TitleSize 14 -XAxisTitleText "Date" -XAxisTitleBold -XAxisTitleSize 12 -YAxisTitleText "KBps" -YAxisTitleBold -YAxisTitleSize 12 -SeriesHeader "KBps" -Row 40 -Column 10
    #$Report | Export-Excel -Path "$($_.Name).xlsx"
    $report | Export-Excel -Path "$file" -WorksheetName "$($_.Name)" -AutoNameRange -ExcelChartDefinition $chart_cpu,$chart_mem,$chart_net
}