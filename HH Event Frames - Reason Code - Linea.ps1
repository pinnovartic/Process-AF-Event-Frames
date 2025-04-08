#region Execution Time
$TimeQuerySt = ConvertFrom-AFRelativeTime -RelativeTime "7-apr"
Write-Host "Fecha Inicio: " $TimeQuerySt
#$TimeQuerySt = $TimeQuerySt.ToLocalTime() 
$TimeQueryEt = ConvertFrom-AFRelativeTime -RelativeTime "8-apr"
Write-Host "Fecha Fin: " $TimeQueryEt
#$TimeQueryEt = $TimeQueryEt.ToLocalTime()
$HH = ($TimeQueryEt - $TimeQuerySt).TotalHours
#endregion
$Linea = "Linea RE010"
$Activos = @("270AG070", "270BB130", "270BB131")

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
    $MyEventFramesOverlapped = Find-AFEventFrame -StartTime ($TimeQuerySt) -EndTime ($TimeQueryEt) -AFSearchMode Overlapped -MaxCount 1000 -ReferencedElementNameFilter $Linea -AFDatabase $AFDB
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
        Write-Host "Evento Linea: " $Event.Name
        Write-Host "Inicio Evento Linea: " $EventST
        Write-Host "Termino Evento Linea: " $EventET

        $EventDuration = ($EventET - $EventST).TotalHours

        foreach ($Activo in $Activos){
            $HRE = 0
            $HPO = 0
            $HMT = 0
            $HEX = 0
            $MyEventFramesOverlapped_Activos = Find-AFEventFrame -StartTime ($EventST) -EndTime ($EventET) -AFSearchMode Overlapped -MaxCount 1000 -ReferencedElementNameFilter $Activo -AFDatabase $AFDB
            foreach($Event_Activo in $MyEventFramesOverlapped_Activos){
                
                Write-Host $Event_Activo.Name
                $EventAttributes_Activo = $Event_Activo.Attributes
                $EventElement_Activo = $Event_Activo.PrimaryReferencedElement
                $EventElementName_Activo = $EventElement_Activo.Name
                $EventST_Activo = $Event_Activo.StartTime
                $EventET_Activo = $Event_Activo.EndTime
                #Evento cruza limites de ventana de revisión
                If ($EventST_Activo.LocalTime -lt $EventST.LocalTime){
                    $EventST_Activo = $EventST
                }
                If ($EventET_Activo.LocalTime -gt $EventET.LocalTime){
                        $EventET_Activo = $EventST
                }
                #Write-Host "Inicio Evento Activo: " $EventST_Activo
                #Write-Host "Termino Evento Activo: " $EventET_Activo

                $EventReason_Activo = $EventAttributes_Activo.Item("Reason")
                If ($EventReason_Activo -ne $null){
                    $EventReasonValue_Activo = $EventReason_Activo.GetValue($EventST_Activo).Value.ShortName
                    If ($EventReasonValue_Activo -ne ""){
                        Write-Host "Reason: " $EventReasonValue_Activo
                    }                            
                    switch ($EventReasonValue_Activo) {
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
        write-Host "HPO:" $HPO
        Write-Host "HMT:" $HMT
        Write-Host "HEX:" $HEX
        
    }
    $TimeQuerySt = $TimeQuerySt.AddDays(1)
    $TimeQueryEt = $TimeQueryEt.AddDays(1)
}