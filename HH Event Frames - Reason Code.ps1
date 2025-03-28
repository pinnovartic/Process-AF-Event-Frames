#region Execution Time
$TimeQuerySt = ConvertFrom-AFRelativeTime -RelativeTime "4-feb"
Write-Host "Fecha Inicio: " $TimeQuerySt
#$TimeQuerySt = $TimeQuerySt.ToLocalTime() 
$TimeQueryEt = ConvertFrom-AFRelativeTime -RelativeTime "5-feb"
Write-Host "Fecha Fin: " $TimeQueryEt
#$TimeQueryEt = $TimeQueryEt.ToLocalTime()
$HH = ($TimeQueryEt - $TimeQuerySt).TotalHours
#endregion
$Activos = @("130HV201", "150TT001")


#region Config read
$Str_AFServer = "CNECLPI03"
$Str_AFUser = "CNECLPI03\Administrador"
$Str_AFPassword = "Sgtm2020"
$Str_AFDatabase = "Ecometales"
#endregion

#region PI AF Connection
$secure_pass = ConvertTo-SecureString -String $Str_AFPassword -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential ($Str_AFUser, $secure_pass)
    try{
        $AFServer = Get-AFServer $Str_AFServer
        $AF_Connection = Connect-AFServer -WindowsCredential $credentials -AFServer $AFServer
        $AFDB = Get-AFDatabase -Name $Str_AFDatabase -AFServer $AFServer
        $AFRootElement = Get-AFElement -AFDatabase $AFDB -Name "Estado Equipos"
    }catch { 
        $e = $_.Exception
        $msg = $e.Message
        while ($e.InnerException) {
            $e = $e.InnerException
            $msg += "`n" + $e.Message
            }
        $msg
    }
#endregion

#region PI AF Event Frames

$CurrentDate = Get-Date

While ($TimeQueryEt.ToLocalTime() -lt $CurrentDate){

    $MyEventFramesOverlapped = Find-AFEventFrame -StartTime ($TimeQuerySt) -EndTime ($TimeQueryEt) -AFSearchMode Overlapped -MaxCount 1000 -AFDatabase $AFDB

    foreach ($Activo in $Activos){
        $HRE = 0
        $HPO = 0
        $HMT = 0
        $HEX = 0
        foreach($Event in $MyEventFramesOverlapped){
            $EventAttributes = $Event.Attributes
            $EventElement = $Event.PrimaryReferencedElement
            $EventElementName = $EventElement.Name
            $EventST = $Event.StartTime
            $EventET = $Event.EndTime

            #Evento cruza limites de ventana de revisión
            If ($EventST.LocalTime -lt $TimeQuerySt.ToLocalTime()){
                $EventST = $TimeQuerySt
            }
            If ($EventET.LocalTime -gt $TimeQueryEt.ToLocalTime()){
                 $EventET = $TimeQueryEt
            }

            $EventDuration = ($EventET - $EventST).TotalHours
            $EventReason = $EventAttributes.Item("Reason")
            If ($EventReason -ne $null){
                If ($Activo -eq $EventElementName){
                    $EventReasonValue = $EventReason.GetValue($EventST).Value.ShortName
                    switch ($EventReasonValue) {
                        "Reservas" {$HRE = $HRE + $EventDuration}
                        "Pérdida Operacional" {$HPO = $HPO + $EventDuration }
                        "Preventivo" {$HMT = $HMT + $EventDuration}
                        "Correctivo" {$HMT = $HMT + $EventDuration}
                        "Horario Inhábil" {$HEX = $HEX + $EventDuration}
                    }
                }
            }
        }
        Write-Host "Activo:" $Activo
        Write-Host "HRE:" $HRE
        Write-Host "HPO:" $HPO
        Write-Host "HMT:" $HMT
        Write-Host "HEX:" $HEX
    
    #region Definición KPIs

    # Utilización = (HH - HMT - HRE - HPO - HIN)/(HH - HIN)
    # HRE (Horas de Reserva - Clasificado en Evento)
    # HPO (Horas de Pérdida Operacional - Clasificado en Evento)

    # Disponibilidad = 100*($HH - $HMT - $HEX)/($HH - $HEX)
    # HH Horas Hábiles (Dentro del rango)
    # HMT (Suma de Horas Mantenimiento Prev y Corr - Clasificado en Evento)
    # HEX (Horas Externas - Clasificado en Evento)

    #endregion

        #Disponibilidad
        $Disponibilidad = 100*($HH - $HMT - $HEX)/($HH - $HEX)
        Write-Host "Disponibilidad:" $Disponibilidad "%"

        #Utilizacion
        $Utilizacion = 100*($HH - $HMT - $HRE - $HPO - $HEX)/($HH - $HEX)
        Write-Host "Utilizacion:" $Utilizacion "%"

        foreach ($Elem in $AFRootElement.Elements){
            If ($Elem.Name -eq $Activo){
            
                $ElementAttrTiempoExternas = Get-AFAttribute -AFElement $Elem -Name "Tiempo Externas"            
                Set-Variable -Name AFValue_TiempoExternas -Value (New-Object 'OSIsoft.AF.Asset.AFValue')
                $AFValue_TiempoExternas.Timestamp = $TimeQuerySt
                $AFValue_TiempoExternas.Value = $HEX        
                $ElementAttrTiempoExternas.SetValue($AFValue_TiempoExternas)

                $ElementAttrTiempoMantenimiento = Get-AFAttribute -AFElement $Elem -Name "Tiempo Mantenimiento"            
                Set-Variable -Name AFValue_TiempoMantenimiento -Value (New-Object 'OSIsoft.AF.Asset.AFValue')
                $AFValue_TiempoMantenimiento.Timestamp = $TimeQuerySt
                $AFValue_TiempoMantenimiento.Value = $HMT        
                $ElementAttrTiempoMantenimiento.SetValue($AFValue_TiempoMantenimiento)

                $ElementAttrTiempoPerdidaOperacional = Get-AFAttribute -AFElement $Elem -Name "Tiempo Perdida Operacional"            
                Set-Variable -Name AFValue_TiempoPerdidaOperacional -Value (New-Object 'OSIsoft.AF.Asset.AFValue')
                $AFValue_TiempoPerdidaOperacional.Timestamp = $TimeQuerySt
                $AFValue_TiempoPerdidaOperacional.Value = $HPO        
                $ElementAttrTiempoPerdidaOperacional.SetValue($AFValue_TiempoPerdidaOperacional)

                $ElementAttrTiempoReservas = Get-AFAttribute -AFElement $Elem -Name "Tiempo Reservas"            
                Set-Variable -Name AFValue_TiempoReservas -Value (New-Object 'OSIsoft.AF.Asset.AFValue')
                $AFValue_TiempoReservas.Timestamp = $TimeQuerySt
                $AFValue_TiempoReservas.Value = $HRE        
                $ElementAttrTiempoReservas.SetValue($AFValue_TiempoReservas)
            }
        }      
    }
    $TimeQuerySt = $TimeQuerySt.AddDays(1)
    $TimeQueryEt = $TimeQueryEt.AddDays(1)
}





