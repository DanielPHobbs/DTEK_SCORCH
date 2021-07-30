########################################################################################
# SMA Runbook Toolkit (SMART) Visualization Tool
# Version 1.0
# Published on Building Clouds blog : http://aka.ms/buildingclouds
# by the Windows Server and System Center CAT team
# Please send feedback to brunosa@microsoft.com
########################################################################################
# Special thanks to Andrew Luty from the Windows Azure Pack (WAP) team,
# For his help with the DGML output scripts
########################################################################################

    param (
    [String]$WSEndPoint = "https://localhost",
    [String]$WSPort = "9090",
    [Bool]$RunbooksCacheOptimization = $False
    )

$ToolVersion = "1.0"

########################################################################################
#Functions
########################################################################################


function ClearNode()

{
 
    param (
        [System.Windows.Controls.TreeViewItem]$TreeNode
    )
       ForEach ($NewNode In $TreeNode.Items)
              {
              $NewNode.IsSelected = $false
              ClearNode($NewNode)
              }
}  

function ExportRoutine()
{

    param (
    [PSObject[]]$RunbooksArray
    )


    write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Computing Runbooks Details and Dependencies...(Step 2 of 3)"
    $DetailedRunbooksArray = QueryRunbookStatistics($RunbooksArray)
    write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Exporting content to the chosen file formats...(Step 3 of 3)"

     $ExportFileName = $FORM.FindName('ExportName').Text
     If ($FORM.FindName('ExportPSConsole').IsChecked)
            {
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] PowerShell Console Export Started..."
            Foreach ($Runbook in $DetailedRunbooksArray)
            
                {
                If (($Runbook.SubRunbooks.Keys.Count) -ne 0)
                    {
                    write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Runbook dependencies for" $Runbook.RunbookName
                    foreach ($Key in $Runbook.SubRunbooks.Keys) {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] - "$Key "( x"($Runbook.SubRunbooks[$Key].Split("-").Count)")" "( Lines" $Runbook.SubRunbooks[$Key]")"}
                    }
                If (($Runbook.Variables.Keys.Count) -ne 0)
                    {
                    write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Variables dependencies for" $Runbook.RunbookName
                    foreach ($Key in $Runbook.Variables.Keys) {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] - "$Key "( x"($Runbook.Variables[$Key].Split("-").Count)")" "( Lines" $Runbook.Variables[$Key]")"}
                    }
                If (($Runbook.Credentials.Keys.Count) -ne 0)
                    {
                    write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Credentials dependencies for" $Runbook.RunbookName
                    foreach ($Key in $Runbook.Credentials.Keys) {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] - "$Key "( x"($Runbook.Credentials[$Key].Split("-").Count)")" "( Lines" $Runbook.Credentials[$Key]")"}
                    }
                If (($Runbook.Certificates.Keys.Count) -ne 0)
                    {
                    write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Certificates dependencies for" $Runbook.RunbookName
                    foreach ($Key in $Runbook.Certificates.Keys) {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] - "$Key "( x"($Runbook.Certificates[$Key].Split("-").Count)")" "( Lines" $Runbook.Certificates[$Key]")"}
                    }
                If (($Runbook.Connections.Keys.Count) -ne 0)
                    {
                    write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Connections dependencies for" $Runbook.RunbookName
                    foreach ($Key in $Runbook.Connections.Keys) {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] - "$Key "( x"($Runbook.Connections[$Key].Split("-").Count)")" "( Lines" $Runbook.Connections[$Key]")"}
                    }
                }
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] PowerShell Console Export Finished"
            } else {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] PowerShell Console Export not requested"}

    If ($FORM.FindName('ExportVSDCB').IsChecked)
            {
            If (Test-Path -Path $Global:VisioTemplate)
                {

                $NbRows = [Math]::floor([math]::sqrt($DetailedRunbooksArray.Count))
                If ($NbRows -eq 0) { $NbRows = 1}

                write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Visio Export Started..."

                $ListShapes =@()
                $ListShapes.Clear()
                $ListShapesName =@()
                $ListShapesName.Clear()

                write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Creating header"
                $vApp = New-Object -ComObject Visio.Application
                $vDoc = $vApp.Documents.Add("")
                $vStencil = $vApp.Documents.OpenEx($VisioTemplate, 4)
                $SpecificvStencil = $vStencil.Masters | where-object {$_.Name -eq "Process"}
        
                #Size the page dynamically
                $vApp.ActivePage.AutoSize = $true

                #Add a title
                $vToShape = $vApp.ActivePage.Drop($SpecificvStencil, 4, 10)
                #$vStencil.Masters.Name
                $vToShape.Text = "Runbook Dependency Graph"
                $vToShape.Cells("Char.Size").Formula = "= 30 pt."
                $vToShape.Cells("Width").Formula = "= 7"

                #Find and draw subrunbooks
                write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Drawing activities"
                $vFlowChartMaster = $vStencil.Masters | where-object {$_.Name -eq $Global:VisioStencil}
                Foreach ($Runbook in $DetailedRunbooksArray)
                    {
                    write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Visio: Working with" $Runbook.RunbookName
                    $vApp.ActiveWindow.DeselectAll()
                    $vToShape = $vApp.ActivePage.Drop($vFlowChartMaster, ($ListShapes.Count % $NbRows)*2, [Math]::floor($ListShapes.Count / $NbRows)*2)
                    $vToShape.Text = $Runbook.RunbookName
                    $vToShape.Cells("Para.HorzAlign").Formula = "=2"
                    $vToShape.Cells("LeftMargin").Formula = "=0.5"
                    $vSel = $vApp.ActiveWindow.Selection
                    If ($FORM.FindName('GroupThumbnails').IsChecked -eq $True)
                        {$vSel.Group()}
                    $ListShapes += $vToShape
                    $ListShapesName += $Runbook.RunbookName
                    $vSel.DeselectAll()
                    }
                #Find and draw links to subrunbooks, and build variables output
                #http://msdn.microsoft.com/en-us/library/office/aa221686(v=office.11).aspx
                write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Drawing links"
                $vConnectorMaster = $vStencil.Masters | where-object {$_.Name -eq "Dynamic Connector"}
                $VariablesOuput = "VARIABLES DEPENDENCIES:`r`n"
                $CredentialsOuput = "CREDENTIALS DEPENDENCIES:`r`n"
                $CertificatesOuput = "CERTIFICATES DEPENDENCIES:`r`n"
                $ConnectionsOuput = "CONNECTIONS DEPENDENCIES:`r`n"
                Foreach ($Runbook in $DetailedRunbooksArray)
                    {
                    foreach ($Key in $Runbook.SubRunbooks.Keys)
                        {
                        If ($ListShapesName.IndexOf(($Global:FullRunbookList | Where-Object {$_.RunbookName -eq $Key}).RunbookName) -eq -1)
                            #There is a child runbook that is not part of the analyzed scope
                            {
                            write-host  -ForegroundColor red "["(date -format "HH:mm:ss")"] -- WARNING : Child runbook" (($Global:FullRunbookList | Where-Object {$_.RunbookName -eq $Key}).RunbookName) "is not part of analyzed scope and may require additional tagging. It will however be correctly represented in the Visio diagram, with partial information"
                            $vApp.ActiveWindow.DeselectAll()
                            $vToShape = $vApp.ActivePage.Drop($vFlowChartMaster, ($ListShapes.Count % $NbRows)*2, [Math]::floor($ListShapes.Count / $NbRows)*2)
                            $vToShape.Text = (($Global:FullRunbookList | Where-Object {$_.RunbookName -eq $Key}).RunbookName)
                            $vToShape.Cells("Para.HorzAlign").Formula = "=2"
                            $vToShape.Cells("LeftMargin").Formula = "=0.5"
                            Try {$vToShape.CellsSRC(1,3,0).FormulaU = "RGB(255,0,0)"}
                                Catch {$vToShape.CellsSRC(1,3,0).FormulaU = "RGB(255,0,0)"}
                            $vSel = $vApp.ActiveWindow.Selection
                            If ($FORM.FindName('GroupThumbnails').IsChecked -eq $True)
                                {$vSel.Group()}
                            $ListShapes += $vToShape
                            $ListShapesName += (($Global:FullRunbookList | Where-Object {$_.RunbookName -eq $Key}).RunbookName)
                            $vSel.DeselectAll()
                            }
                        $vConnector = $vApp.ActivePage.Drop($vConnectorMaster, 0, 0)
                        $vConnector.Cells("EndArrow").Formula = "=4"
                        $vBeginCell = $vConnector.Cells("BeginX")
                        $vFromShape = $ListShapes.Item($ListShapesName.IndexOf($Runbook.RunbookName))
                        $vBeginCell.GlueTo($vFromShape.Cells("Align" + $Global:VisioGlueFrom))
                        $vEndCell = $vConnector.Cells("EndX")
                        $vConnector.Cells("ConLineRouteExt").ResultIUForce = 2
                        $vToShape = $ListShapes.Item($ListShapesName.IndexOf(($Global:FullRunbookList | Where-Object {$_.RunbookName -eq $Key}).RunbookName))
                        $vEndCell.GlueTo($vToShape.Cells("Align" + $Global:VisioGlueTo))
                        $vConnector.Text = $Runbook.SubRunbooks[$Key]
                        $vConnector.SendToBack()
                        }
                    If (($Runbook.Variables.Keys.Count) -ne 0)
                        {
                        $VariablesOuput = $VariablesOuput + "Variables dependencies for " + $Runbook.RunbookName + "`r`n"
                        foreach ($Key in $Runbook.Variables.Keys)
                            {
                            $VariablesOuput = $VariablesOuput + " - " + $Key + "( x" + ($Runbook.Variables[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Variables[$Key] + ")`r`n"
                            }
                        }
                    If (($Runbook.Credentials.Keys.Count) -ne 0)
                        {
                        $CredentialsOuput = $CredentialsOuput + "Credentials dependencies for " + $Runbook.RunbookName + "`r`n"
                        foreach ($Key in $Runbook.Credentials.Keys)
                            {
                            $CredentialsOuput = $CredentialsOuput + " - " + $Key + "( x" + ($Runbook.Credentials[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Credentials[$Key] + ")`r`n"
                            }
                        }
                    If (($Runbook.Certificates.Keys.Count) -ne 0)
                        {
                        $CertificatesOuput = $CertificatesOuput + "Certificates dependencies for " + $Runbook.RunbookName + "`r`n"
                        foreach ($Key in $Runbook.Certificates.Keys)
                            {
                            $CertificatesOuput = $CertificatesOuput + " - " + $Key + "( x" + ($Runbook.Certificates[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Certificates[$Key] + ")`r`n"
                            }
                        }
                    If (($Runbook.Connections.Keys.Count) -ne 0)
                        {
                        $ConnectionsOuput = $ConnectionsOuput + "Connections dependencies for " + $Runbook.RunbookName + "`r`n"
                        foreach ($Key in $Runbook.Connections.Keys)
                            {
                            $ConnectionsOuput = $ConnectionsOuput + " - " + $Key + "( x" + ($Runbook.Connections[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Connections[$Key] + ")`r`n"
                            }
                        }
                    }
                write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Summarizing variables dependencies"
                $vShapeAssets = $vApp.ActivePage.Drop($SpecificvStencil, 12, 3)
                $vShapeAssets.Text = $VariablesOuput + $CredentialsOuput + $CertificatesOuput + $ConnectionsOuput
                $vShapeAssets.Cells("Char.Size").Formula = "= 12 pt."
                $vShapeAssets.Cells("Width").Formula = "= 7"
                $visSectionParagraph = 4
                $visHorzAlign = 6
                $vShapeAssets.CellsSRC($visSectionParagraph, 0, $visHorzAlign).FormulaU = "0"  
                If ($FORM.FindName('SaveAndCloseVSD').IsChecked -eq $True)
                    {        
                    write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Exporting to Visio file $ExportFileName.VSDX"
                    $vDoc.SaveAs((Get-Location -PSProvider FileSystem).ProviderPath + "\" + $ExportFileName + ".VSDX")
                    $vDoc.Close()
                    $vApp.Quit()
                    }
                write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Visio Export Finished"
                }
                else
                {
                write-host  -ForegroundColor red "["(date -format "HH:mm:ss")"] Visio Templates path is not valid. The file path does not resolve. Please update the Visio settings with a valid path and retry..."
                }
     } else {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Visio Export not requested"}

     If ($FORM.FindName('ExportVSCB').IsChecked)
            {
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Visual Studio Export Started..."
            $xml = New-Graph
            Set-GraphTitle -xml $xml -title "Workflow dependencies"
            $nodecount = 0
            $ListNodesName =@()
            $ListNodesName.Clear()
            Foreach ($Runbook in $DetailedRunbooksArray)
                {
                    write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Visual Studio: Working with" $Runbook.RunbookName
                    $NewNode = Add-GraphNode -xml $xml -id $Runbook.RunbookName -label $Runbook.RunbookName -category 'Workflow' -group 'Collapsed' # 'Expanded'
                    $NewNode.SetAttribute("Description", $Runbook.RunbookDescription)     
                    $NewNode.SetAttribute("Tags", $Runbook.RunbookTags)
                    If (($Runbook.Variables.Keys.Count) -ne 0)
                        {
                        $VariablesNodeOutput= ""
                        foreach ($Key in $Runbook.Variables.Keys)
                            {$VariablesNodeOutput += $Key + "( x" + ($Runbook.Variables[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Variables[$Key] + ");"}
                        $NewNode.SetAttribute("Variable", $VariablesNodeOutput.Substring(0, $VariablesNodeOutput.Length-1))
                        }
                    If (($Runbook.Credentials.Keys.Count) -ne 0)
                        {
                        $CredentialsNodeOutput= ""
                        foreach ($Key in $Runbook.Credentials.Keys)
                            {$CredentialsNodeOutput += $Key + "( x" + ($Runbook.Credentials[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Credentials[$Key] + "):"}
                        $NewNode.SetAttribute("Credential", $CredentialsNodeOutput.Substring(0, $CredentialsNodeOutput.Length-1))
                        }
                    If (($Runbook.Certificates.Keys.Count) -ne 0)
                        {
                        $CertificatesNodeOutput= ""
                        foreach ($Key in $Runbook.Certificates.Keys)
                            {$CertificatesNodeOutput +=  $Key + "( x" + ($Runbook.Certificates[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Certificates[$Key] + ");"}
                        $NewNode.SetAttribute("Certificate", $CertificatesNodeOutput.Substring(0, $CertificatesNodeOutput.Length-1))
                        }
                    If (($Runbook.Connections.Keys.Count) -ne 0)
                        {
                        $ConnectionsNodeOutput= ""
                        foreach ($Key in $Runbook.Connections.Keys)
                            {$ConnectionsNodeOutput +=  $Key + "( x" + ($Runbook.Connections[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Connections[$Key] + ");"}
                        $NewNode.SetAttribute("Connection", $ConnectionsNodeOutput.Substring(0, $ConnectionsNodeOutput.Length-1))
                        }
                    $ListNodesName+= $Runbook.RunbookName
                }
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Drawing links"
            Foreach ($Runbook in $DetailedRunbooksArray)
                {
                    foreach ($Key in $Runbook.SubRunbooks.Keys)
                        {
                        if ($ListNodesName -notcontains $Key)
                            #There is a child runbook that is not part of the analyzed scope
                            {
                            write-host  -ForegroundColor red "["(date -format "HH:mm:ss")"] -- WARNING : Child runbook" (($Global:FullRunbookList | Where-Object {$_.RunbookName -eq $Key}).RunbookName) "is not part of analyzed scope and may require additional tagging. It will however be correctly represented in the Visio diagram, with partial information"
                            $NewNode = Add-GraphNode -xml $xml -id $Key -label $Key -category 'Workflow-nonscoped' -group 'Collapsed' # 'Expanded'
                            $NewNode.SetAttribute("Description", (($Global:FullRunbookList | Where-Object {$_.RunbookName -eq $Key}).Description))
                            $NewNode.SetAttribute("Tags", (($Global:FullRunbookList | Where-Object {$_.RunbookName -eq $Key}).Tags))
                            $ListNodesName+= $Runbook.RunbookName
                            }                        
                        Add-GraphLink -xml $xml -source $Runbook.RunbookName -target $Key -label $Runbook.SubRunbooks[$Key] -category 'contains'
                        }
                }                
            $xml.Save((Get-Location -PSProvider FileSystem).ProviderPath + "\" + $ExportFileName + ".dgml")
            If ($FORM.FindName('SaveAndCloseVS').IsChecked -eq $False)
                {  
                Invoke-Item ((Get-Location -PSProvider FileSystem).ProviderPath + "\" + $ExportFileName + ".dgml")
                }
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Visual Studio Export Finished"
            } else {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Visual Studio not requested"}

     If ($FORM.FindName('ExportDOCCB').IsChecked)
            {
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Word Export Started..."
            $oWord = New-Object -ComObject Word.Application
            $oWord.Visible = $True
            $oDoc = $oWord.Documents.Add()
            $oPara1 = $oDoc.Content.Paragraphs.Add()
            $oPara1.Range.Text = "Runbook Dependencies Details"
            $oPara1.Range.Font.Bold = $True
            $oPara1.Range.Font.Size = 28
            $oPara1.Format.SpaceAfter = 24    #24 pt spacing after paragraph.
            $oPara1.Range.InsertParagraphAfter()
            $oTable = $oDoc.Tables.Add($oDoc.Bookmarks.Item("\endofdoc").Range, 1, 5)
            $oTable.Range.ParagraphFormat.SpaceAfter = 6
            $oTable.Range.Font.Size = 8
            $oTable.Range.Font.Bold = $True
            $oTable.Range.Borders.Enable = $True
            $oTable.Range.Borders.OutsideLineStyle = 7
            $oTable.Range.Borders.InsideLineStyle = 0
            $oTable.Cell(1, 1).Range.Text = "Runbook"
            $oTable.Cell(1, 2).Range.Text = "Tags"
            $oTable.Cell(1, 3).Range.Text = "Description"
            $oTable.Cell(1, 4).Range.Text = "Runbooks Depedencies"
            $oTable.Cell(1, 5).Range.Text = "Assets dependencies"
            $oTable.Columns.Item(4).Width = $oWord.InchesToPoints(2)
            $oTable.Columns.Item(2).Width = $oWord.InchesToPoints(1)
            $oTable.Columns.Item(5).Width = $oWord.InchesToPoints(1)
            $oTable.Columns.Item(3).Width = $oWord.InchesToPoints(1)
            $r = 2
            Foreach ($Runbook in $DetailedRunbooksArray)
                {
                write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Word: Working with" $Runbook.RunbookName
                $oTable.Rows.Add()
                $oTable.Rows.Item($r).Range.Font.Bold = $False
                $oTable.Rows.Item($r).Range.Borders.Enable = $True
                $oTable.Rows.Item($r).Range.Borders.OutsideLineStyle = 1
                $oTable.Rows.Item($r).Range.Borders.InsideLineStyle = 0
                $oPara1 = $oTable.Cell($r, 1).Range.Paragraphs.Add()
                $oPara1.Range.Text = $Runbook.RunbookName
                $oPara1 = $oTable.Cell($r, 2).Range.Paragraphs.Add()
                $oPara1.Range.Text = $Runbook.RunbookTags
                $oPara1 = $oTable.Cell($r, 3).Range.Paragraphs.Add()
                $oPara1.Range.Text = $Runbook.RunbookDescription
                foreach ($Key in $Runbook.SubRunbooks.Keys)
                    {
                    $oPara1 = $oTable.Cell($r, 4).Range.Paragraphs.Add()
                    $oPara1.Range.Text = $Key + "( x" + ($Runbook.SubRunbooks[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.SubRunbooks[$Key] + ")"
                    }
                If (($Runbook.Variables.Keys.Count) -ne 0)
                    {
                    $oPara1 = $oTable.Cell($r, 5).Range.Paragraphs.Add()
                    $oPara1.Range.Text = "VARIABLES"
                    foreach ($Key in $Runbook.Variables.Keys)
                        {
                        $oPara1 = $oTable.Cell($r, 5).Range.Paragraphs.Add()
                        $oPara1.Range.Text = $Key + "( x" + ($Runbook.Variables[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Variables[$Key] + ")"
                        }
                    }
                If (($Runbook.Credentials.Keys.Count) -ne 0)
                    {
                    $oPara1 = $oTable.Cell($r, 5).Range.Paragraphs.Add()
                    $oPara1.Range.Text = "CREDENTIALS"
                    foreach ($Key in $Runbook.Credentials.Keys)
                        {
                        $oPara1 = $oTable.Cell($r, 5).Range.Paragraphs.Add()
                        $oPara1.Range.Text = $Key + "( x" + ($Runbook.Credentials[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Credentials[$Key] + ")"
                        }
                    }
                If (($Runbook.Certificates.Keys.Count) -ne 0)
                    {
                    $oPara1 = $oTable.Cell($r, 5).Range.Paragraphs.Add()
                    $oPara1.Range.Text = "CERTIFICATES"
                    foreach ($Key in $Runbook.Certificates.Keys)
                        {
                        $oPara1 = $oTable.Cell($r, 5).Range.Paragraphs.Add()
                        $oPara1.Range.Text = $Key + "( x" + ($Runbook.Certificates[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Certificates[$Key] + ")"
                        }
                    }
                If (($Runbook.Connections.Keys.Count) -ne 0)
                    {
                    $oPara1 = $oTable.Cell($r, 5).Range.Paragraphs.Add()
                    $oPara1.Range.Text = "CONNECTIONS"
                    foreach ($Key in $Runbook.Connections.Keys)
                        {
                        $oPara1 = $oTable.Cell($r, 5).Range.Paragraphs.Add()
                        $oPara1.Range.Text = $Key + "( x" + ($Runbook.Connections[$Key].Split("-").Count) + ")" + "( Lines" + $Runbook.Connections[$Key] + ")"
                        }
                    }
                $r = $r + 1
                }
            If ($FORM.FindName('SaveAndCloseDOC').IsChecked -eq $True)
                {        
                write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Exporting to Word file $ExportFileName.DOCX"
                $oDoc.SaveAs([ref]((Get-Location -PSProvider FileSystem).ProviderPath + "\" + $ExportFileName + ".DOCX"))
                $oDoc.Close()
                $oWord.Quit()
                }
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Word Export Finished"
            } else {write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] Word Export not requested"}

     write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Runbooks export finished."
     popup2 -Message ("Runbooks export finished") -ClosedExternally $False -NbLines 2
}

function CacheFullRunbookList()
{

    $J = Start-Job -ScriptBlock {
            param ($WS,$Port)
            get-smarunbook -WebServiceEndpoint $WS -Port $Port | select RunbookName, RunbookID, Tags, Description
            } -ArgumentList $Global:SMAWSEndPoint, $Global:SMAWSPort
    $Loop = 1
    while ( ($J.state -ne "Completed") -And ($Loop -lt 40) )
            {
            Start-Sleep(3)
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Job is still running..."
            $Loop = $Loop+1
            }
    If ($J.State -eq "Completed")
            {
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Job completed. If there was an error it should be displayed below."
            $Global:FullRunbookList  = Receive-Job -Job $J -Keep
            }
}

function QueryRunbookStatistics()
{
param (
    [PSObject[]]$Runbooks
)

$Count = 0
$Detailedrunbooks = @()


write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Querying runbooks definitions...This may take from 5 to 30 sec, depending on the number of runbooks..."
#Adding a 'try" clause for the RunbookIDs set to 'LOCALFILE'. Also, this syntax - instead of a where clause - makes sure we return all Runbook definitions, and not just the first 100.
If (($Runbooks | ? RunbookID -notmatch "LOCALFILE*").Count -ne 0)
    {
    $RunbookBodyList = ($Runbooks | ? RunbookID -notmatch "LOCALFILE*").RunbookID | get-smarunbookdefinition  -WebServiceEndpoint  $Global:SMAWSEndPoint -Port $Global:SMAWSPort -type published | select RunbookVersion, Content
    }

write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Querying runbooks definitions...Done!"
write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Querying runbooks details...If this takes too long, please look at the performance tips on the blog post"    
Write-Progress -Activity "Querying runbooks details..." -PercentComplete (100*$Count/($Runbooks.Count))

foreach ($Runbook in $Runbooks)
    {    
    $SubRunbooksList = @{}
    $VariablesList = @{}
    $CredentialsList = @{}
    $ConnectionsList = @{}
    $CertificatesList = @{}
    switch ($Runbook.RunbookID)
        {
        "LOCALFILE" {$RunbookBody = Get-Content -raw -path ($FORM.FindName('PickFolderTextBoxName').Text + "\" + $Runbook.RunbookName + ".ps1")}
        "LOCALFILE-LRO" {$RunbookBody = Get-Content -raw -path ($RunbooksCacheDirectory + $Runbook.RunbookName + ".ps1")}
        default {$RunbookBody = ($RunbookBodyList | Where-Object {$_.RunbookVersion.RunbookID -eq $Runbook.RunbookID}).Content}
        }
    $AbtractSyntaxTree = [System.Management.Automation.Language.Parser]::ParseInput($RunbookBody, [ref]$null, [ref]$null)
    $eap = $ErrorActionPreference = "SilentlyContinue"
    $AbtractSyntaxTree.FindAll({$args[0] -is [System.Management.Automation.Language.CommandAst]},$true) | foreach { 
        $Command = $_.CommandElements[0]
        #write-host $Command
        if ($Alias = $Runbooks | ? RunbookID -match "LOCALFILE*" | ? RunbookName -eq $Command)
            #we track subrunbook calls and compare with runbooks in local files, if any
            {
            #write-host "found match in local files"
            If ($SubRunbooksList.Contains($Alias.RunbookName))
                {
                    $SubRunbooksList[$Alias.RunbookName] = $SubRunbooksList[$Alias.RunbookName] + "-" + [string]$Command.Extent.StartLineNumber
                }
                else
                {
                    $SubRunbooksList.add($Alias.RunbookName, [string]$Command.Extent.StartLineNumber)
                }
            }
            else
            {
            if ($Global:FullRunbookList.RunbookName -contains $Command.Value)
                #we track subrunbook calls and compare with runbooks online in SMA, if any
                {
                #write-host "found match in SMA"
                If ($SubRunbooksList.Contains($Command.Value))
                    {
                        $SubRunbooksList[$Command.Value] = $SubRunbooksList[$Command.Value] + "-" + [string]$Command.Extent.StartLineNumber
                    }
                    else
                    {
                        $SubRunbooksList.add($Command.Value, [string]$Command.Extent.StartLineNumber)
                    }
                }
            }
        if (($Command.Value -eq "Get-AutomationVariable") -or ($Command.Value -eq "Set-AutomationVariable"))
            #we track variable calls and actual variables used
            {
            If ($VariablesList.Contains(($_.CommandElements[2]).Value.Replace("'","")))
                {
                    $VariablesList[($_.CommandElements[2]).Value.Replace("'","")] = $VariablesList[($_.CommandElements[2]).Value.Replace("'","")] + "-" + [string]$Command.Extent.StartLineNumber
                }
                else
                {
                    $VariablesList.add(($_.CommandElements[2]).Value.Replace("'",""),  [string]$Command.Extent.StartLineNumber)
                }
            }  
        if ($Command.Value -eq "Get-AutomationConnection")
            #we track variable calls and actual connections used
            {
            If ($ConnectionsList.Contains(($_.CommandElements[2]).Value.Replace("'","")))
                {
                    $ConnectionsList[($_.CommandElements[2]).Value.Replace("'","")] = $ConnectionsList[($_.CommandElements[2]).Value.Replace("'","")] + "-" + [string]$Command.Extent.StartLineNumber
                }
                else
                {
                    $ConnectionsList.add(($_.CommandElements[2]).Value.Replace("'",""),  [string]$Command.Extent.StartLineNumber)
                }
            }  
        if ($Command.Value -eq "Get-AutomationCertificate")
            #we track variable calls and actual certificates used
            {
            If ($CertificatesList.Contains(($_.CommandElements[2]).Value.Replace("'","")))
                {
                    $CertificatesList[($_.CommandElements[2]).Value.Replace("'","")] = $CertificatesList[($_.CommandElements[2]).Value.Replace("'","")] + "-" + [string]$Command.Extent.StartLineNumber
                }
                else
                {
                    $CertificatesList.add(($_.CommandElements[2]).Value.Replace("'",""),  [string]$Command.Extent.StartLineNumber)
                }
            }  
        if ($Command.Value -eq "Get-AutomationPSCredential")
            #we track variable calls and actual PS credentials used
            {
            If ($CredentialsList.Contains(($_.CommandElements[2]).Value.Replace("'","")))
                {
                    $CredentialsList[($_.CommandElements[2]).Value.Replace("'","")] = $CredentialsList[($_.CommandElements[2]).Value.Replace("'","")] + "-" + [string]$Command.Extent.StartLineNumber
                }
                else
                {
                    $CredentialsList.add(($_.CommandElements[2]).Value.Replace("'",""),  [string]$Command.Extent.StartLineNumber)
                }
            }  
        }
    $DetailedRunbooks += New-Object PSObject -Property @{ 
          RunbookName = $Runbook.RunbookName
          RunbookID = $Runbook.RunbookID
          RunbookTags = ($Global:FullRunbookList | Where-Object {$_.RunbookName -eq $Runbook.RunbookName}).Tags
          RunbookDescription = ($Global:FullRunbookList | Where-Object {$_.RunbookName -eq $Runbook.RunbookName}).Description
          SubRunbooks = $SubRunbooksList
          Variables = $VariablesList
          Connections = $ConnectionsList
          Certificates = $CertificatesList
          Credentials = $CredentialsList
          } 
    $ErrorActionPreference =$eap
    $Count  += 1
    Write-Progress -Activity "Querying runbooks details..." -PercentComplete (100*$Count/($Runbooks.Count))
    }
Write-Progress -Activity "Querying runbooks details..." -Completed
write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Querying runbooks details..Done!"    
return $DetailedRunbooks
}

function Popup2()
{

param (
    [String]$Message,
    [Boolean]$ClosedExternally,
    [Boolean]$ShowDialog,
    [int]$NbLines
)
        $MB = New-Object System.Windows.Window
        $MB.Width = 500
        $MB.Height = 55*$NbLines
        $MB.WindowStyle = "None"

        $MBLabel = New-Object System.Windows.Controls.Label
        $MBLabel.HorizontalAlignment = "Center"
        $MBLabel.FontFamily = "Verdana"
        $MBLabel.Content = $Message

        $MBStackPanel = New-Object System.Windows.Controls.StackPanel
        $MBStackPanel.VerticalAlignment = "Top"

        If ($ClosedExternally -eq $False){
            $MBButton = New-Object System.Windows.Controls.Button
            $MBButton.Content = "Close"
            $MBButton.HorizontalAlignment = "Center"
            $MBButton.Width="100"
            $MBButton.Add_Click({
            $this.parent.parent.Close()
            })
            $result=$MBStackPanel.Children.Add($MBLabel)
            $result=$MBStackPanel.Children.Add($MBButton)
            $MB.Content = $MBStackPanel
        }
        else
        {
        $MB.Content = $MBLabel
        }
        
        If ($ShowDialog){$MB.ShowDialog()} else {$MB.Show()}
        #write-host ("POPUP : " + $Message.Replace("`r`n"," "))
        Return $MB
        
}


function ListRunbooks()
{

param (
    [System.Windows.Controls.TreeView]$Tree
)

        $Global:ProgressCount = 0

        write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Trying to connect to web service $Global:SMAWSEndPoint"
        Write-Progress -Activity "Connecting to Web Service on $Global:SMAWSEndPoint and retrieving Runbooks..."
        $Tree.Items.Clear()
        $eap = $ErrorActionPreference = "SilentlyContinue"
        $TempVar = (get-smavariable -WebServiceEndpoint $Global:SMAWSEndPoint -Port $Global:SMAWSPort | ? name -like "*")
        if (!$?) {
            $ErrorActionPreference =$eap
            $MB = popup2 -Message ("Runbook hierarchy cannot be displayed.`r`nConnection to web service " + $Global:SMAWSEndPoint + " could not be opened.`r`nPlease configure or check the server name on the next screen and try again.") -ClosedExternally $False -NbLines 2 -ShowDialog $True
            $Tree.IsEnabled= $False
            }
            else{  
            $ErrorActionPreference =$eap
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Connected to web service, starting job to retrieve Runbooks..."
            #$NumberOfRunbooks = (get-smarunbook -WebServiceEndpoint $Global:SMAWSEndPoint -Port $Global:SMAWSPort).count

            $J = Start-Job -ScriptBlock {
                param ($WS,$Port)
                get-smarunbook -WebServiceEndpoint $WS -Port $Port | select RunbookName, RunbookId, Tags | Group-Object Tags
                } -ArgumentList $Global:SMAWSEndPoint, $Global:SMAWSPort
            $Loop = 1
            while ( ($J.state -ne "Completed") -And ($Loop -lt 40) )
                {
                Start-Sleep(3)
                write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Job is still running..."
                $Loop = $Loop+1
                }

            If ($J.State -eq "Completed")
                {
                write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Job completed. If there was an error it should be displayed below"
                $data  = Receive-Job -Job $J -Keep
                # write-host $data.GetType()
                #$data | select Group
                $Global:UnfilteredTagArray = @()
                foreach ($TagGroup in $data)
                    {
                    #Fill Runbooks TreeView
                    $NodeRoot = New-Object System.Windows.Controls.TreeViewItem 
                    $NodeRoot.Header = $TagGroup.Name
                    If ($TagGroup.Name -eq "") {$NodeRoot.Header = "[No Tags]"}
                    $NodeRoot.Name = "Folder"
                    $NodeRoot.Tag = "00000000-0000-0000-0000-000000000000"
                    [void]$Tree.Items.Add($NodeRoot)
                    foreach ($Runbook in $TagGroup.Group)
                        {
                        #write-host $Runbook.RunbookName
                        $NewNode = New-Object System.Windows.Controls.TreeViewItem 
                        $NewNode.Header = $Runbook.RunbookName
                        $NewNode.Tag = $Runbook.RunbookId
                        $NewNode.Name = "Runbook"
                        $NewNode.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Blue)
                        [void]$NodeRoot.Items.Add($NewNode)
                        $Global:UnfilteredTagArray += ($TagGroup.Name -split (","))
                        }
                    $Global:ProgressCount = $Global:ProgressCount + ($TagGroup.Count/$NumberOfRunbooks)*100
                    #Write-Progress -Activity "Connecting to web service $Global:SMAWSEndPoint and retrieving Runbooks..." -PercentComplete $Global:ProgressCount
                    }
                $Tree.IsEnabled= $True
                #Fill Tags Dropdown List
                $Global:UnfilteredTagArray = $Global:UnfilteredTagArray | Select-Object -Unique
                $Global:TagArray = @()
                foreach ($NewTag in $Global:UnfilteredTagArray)
                        {
                        If ($NewTag -eq "") {$NewTag = "[No Tags]"}
                        $tmpObject = select-object -inputobject "" TagChecked, TagName
                        $tmpObject.TagChecked = $false
                        $tmpObject.TagName = $NewTag
                        $Global:TagArray += $tmpObject
                        }
                $FORM.FindName('PickTagsComboBox').ItemsSource = $Global:TagArray
                #$FORM.FindName('PickTagsComboBox').IsDropDownOpen = $true
                #Finalize Runbook Parsing Routine
                write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Runbooks parsing finished..."
                $FORM.FindName('ListVariables').IsEnabled = $True
                $FORM.FindName('ExportAllButton').IsEnabled = $True
                $FORM.FindName('ExportButton').IsEnabled = $True
                $FORM.FindName('ListRunbookDetails').IsEnabled = $True
                Write-Progress -Activity "Connecting to web service $Global:SMAWSEndPoint and retrieving Runbooks..." -Complete
                }
            }
            
}


 function LoadConfig()
 {

  param (
    [String]$ConfigFileLocation

  )

        write-host -ForegroundColor gray "["(date -format "HH:mm:ss")"] Loading Configuration from file $ConfigFileLocation..."
        $XmlReader= New-Object System.Xml.XmlTextReader($ConfigFileLocation)
        While ($XmlReader.Read()){
                              If ($XmlReader.NodeType -eq [System.Xml.XmlNodeType]::Element){
                                   switch ($XmlReader.Name){
                                        "WSEndPoint" {$Global:SMAWSEndpoint=$XmlReader.ReadString()}
                                        "WSPort"{$Global:SMAWSPort=$XmlReader.ReadString()}
                                        "VisioTemplate"{$Global:VisioTemplate=$XmlReader.ReadString()}
                                        "VisioStencil"{$Global:VisioStencil=$XmlReader.ReadString()}
                                        "VisioCallout"{$Global:VisioCallout=$XmlReader.ReadString()}
                                        "VisioGlueFrom"{$Global:VisioGlueFrom=$XmlReader.ReadString()}
                                        "VisioGlueTo"{$Global:VisioGlueTo=$XmlReader.ReadString()}
                                    }
                               }
        }
        $XmlReader.Close()
 }

 function SaveConfig()
 {

 param (
    [String]$ConfigFileLocation,
    [String]$WSEndPoint,
    [String]$WSPort,
    [String]$VTemplate,
    [String]$VStencil,
    [String]$VCallout,
    [String]$VGlueFrom,
    [String]$VGlueTo
 )

 write-host -ForegroundColor gray "["(date -format "HH:mm:ss")"] Saving Configuration to file $ConfigFileLocation..."

 $XMLData = 
@”
<DefaultConfiguration>
<WSEndPoint>$WSEndPoint</WSEndPoint>
<WSPort>$WSPort</WSPort>
<VisioTemplate>$VTemplate</VisioTemplate>
<VisioStencil>$VStencil</VisioStencil>
<VisioCallout>$VCallout</VisioCallout>
<VisioGlueFrom>$VGlueFrom</VisioGlueFrom>
<VisioGlueTo>$VGlueTo</VisioGlueTo>
</DefaultConfiguration>
“@

$XMLData | Out-File $ConfigFileLocation -Force

 }


########################################################################################
#Make sure we run elevated, or relaunch as admin
########################################################################################


$CurrentScriptDirectory = $PSCommandPath.Substring(0,$PSCommandPath.LastIndexOf("\"))
Set-Location $CurrentScriptDirectory


    #Thanks to http://gallery.technet.microsoft.com/scriptcenter/63fd1c0d-da57-4fb4-9645-ea52fc4f1dfb
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") 
        if (-not $IsAdmin)  
        {  
            try 
            {  
                $ScriptToLaunch = $PSCommandPath.Substring(0,$PSCommandPath.LastIndexOf("\")) + "\SMART-VT.ps1"
                $arg = "-file `"$($ScriptToLaunch)`"" 
                write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] WARNING : This script should run with administrative rights - Relaunching the script in elevated mode in 3 seconds..."
                start-sleep 3
                Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'

            } 
            catch 
            { 
                write-host -ForegroundColor red "["(date -format "HH:mm:ss")"] Error : Failed to restart script with administrative rights - please make sure this script is launched elevated."  
                break               
            } 
            exit
        }
        else
        {
        write-host -ForegroundColor gray "["(date -format "HH:mm:ss")"] We are running in elevated mode, we can proceed with launching the tool."
        }

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

##########################################################################################
# Visio Settings GUI and functions
########################################################################################## 

[XML]$XAMLVisioSettings = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        ResizeMode="NoResize"
        Title="Visio Settings" Height="225" Width="870">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
        </Grid.RowDefinitions>
        <Label FontWeight="Bold" Content="Please confirm the following Visio settings, and just close the window when finished. These settings will then be automatically applied." VerticalAlignment="Center" Grid.Row="0"></Label>
        <TextBox Text="C:\Program Files (x86)\Microsoft Office\Office15\Visio Content\1033\BASFLO_U.VSSX" Name="VisioTemplateTextBox" Grid.Row="1" HorizontalAlignment="Left" VerticalAlignment="Center"  Margin="250,0,0,0" Width="500"></TextBox>
        <Label Content="Default stencil for activities" VerticalAlignment="Center" Grid.Row="2"></Label>
        <ComboBox Name="VisioStencilComboBox" IsEditable="False" HorizontalAlignment="Left" VerticalAlignment="Center" Width="150"  Margin="250,0,0,0" Grid.Row="2"></ComboBox>
        <Label Content="Default callout for descriptions" VerticalAlignment="Center" Grid.Row="3"></Label>
        <ComboBox Name="VisioCalloutComboBox" IsEditable="False" HorizontalAlignment="Left" VerticalAlignment="Center" Width="150"  Margin="250,0,0,0" Grid.Row="3"></ComboBox>
        <Label Content="Link from" VerticalAlignment="Center" Grid.Row="4"></Label>
        <ComboBox Name="VisioGlueFromComboBox" IsEditable="False" HorizontalAlignment="Left" VerticalAlignment="Center" Width="150"  Margin="250,0,0,0" Grid.Row="4"></ComboBox>
        <Label Content="Link to" VerticalAlignment="Center" Grid.Row="5"></Label>
        <ComboBox Name="VisioGlueToComboBox" IsEditable="False" HorizontalAlignment="Left" VerticalAlignment="Center" Width="150"  Margin="250,0,0,0" Grid.Row="5"></ComboBox>
    </Grid>
</Window>

'@

##########################################################################################
# Directed Graph Modeling Language (DGML) functions
########################################################################################## 

$Dgml = "http://schemas.microsoft.com/vs/2009/dgml"
$DirectedGraphTemplate = @"
<?xml version="1.0" encoding="utf-8" ?>
<DirectedGraph xmlns="http://schemas.microsoft.com/vs/2009/dgml" Title="" GraphDirection="LeftToRight" Layout="Sugiyama">
  <Nodes>
    <!--Node Id="" /-->
  </Nodes>
  <Links>
    <!--Link Source="" Target="" /-->
  </Links>
  <Categories>
    <Category Id="Workflow" />
    <Category Id="Function" />
    <Category Id="Contains" Label="Contains" Description="Whether the source of the link contains the target object" CanBeDataDriven="False" CanLinkedNodesBeDataDriven="True" IncomingActionLabel="Contained By" IsContainment="True" OutgoingActionLabel="Contains" />
    <Category Id="References" Label="References" CanBeDataDriven="True" CanLinkedNodesBeDataDriven="True" IncomingActionLabel="Referenced By" OutgoingActionLabel="References" />
    <Category Id="Comment" Label="Comment" Description="Represents a user defined comment on the diagram" CanBeDataDriven="True" IsProviderRoot="False" NavigationActionLabel="Comments" />
  </Categories>
  <Properties>
    <Property Id="Bounds" DataType="System.Windows.Rect" />
    <Property Id="CanBeDataDriven" Label="CanBeDataDriven" Description="CanBeDataDriven" DataType="System.Boolean" />
    <Property Id="CanLinkedNodesBeDataDriven" Label="CanLinkedNodesBeDataDriven" Description="CanLinkedNodesBeDataDriven" DataType="System.Boolean" />
    <Property Id="GraphDirection" DataType="Microsoft.VisualStudio.Diagrams.Layout.LayoutOrientation" />
    <Property Id="Group" Label="Group" Description="Display the node as a group" DataType="Microsoft.VisualStudio.GraphModel.GraphGroupStyle" />
    <Property Id="IncomingActionLabel" Label="IncomingActionLabel" Description="IncomingActionLabel" DataType="System.String" />
    <Property Id="IsContainment" DataType="System.Boolean" />
    <Property Id="Label" Label="Label" Description="Displayable label of an Annotatable object" DataType="System.String" />
    <Property Id="OutgoingActionLabel" Label="OutgoingActionLabel" Description="OutgoingActionLabel" DataType="System.String" />
    <Property Id="Title" DataType="System.String" />
    <Property Id="ForwardingAddress" DataType="System.Uri" />
  </Properties>
  <Styles>
    <Style TargetType="Node" GroupLabel="Workflow" ValueLabel="True">
      <Condition Expression="HasCategory('Workflow')" />
      <Setter Property="Icon" Value="pack://application:,,,/Microsoft.VisualStudio.Progression.GraphControl;component/Icons/Script.png" />
    </Style>
    <Style TargetType="Node" GroupLabel="Workflow-nonscoped" ValueLabel="True">
      <Condition Expression="HasCategory('Workflow-nonscoped')" />
      <Setter Property="Icon" Value="pack://application:,,,/Microsoft.VisualStudio.Progression.GraphControl;component/Icons/Script.png" />
      <Setter Property="Background" Value="#FFFF0000" />
    </Style>
    <Style TargetType="Node" GroupLabel="Function" ValueLabel="True">
      <Condition Expression="HasCategory('Function')" />
      <Setter Property="Icon" Value="pack://application:,,,/Microsoft.VisualStudio.Progression.GraphControl;component/Icons/Function.png" />
    </Style>
    <Style TargetType="Node" GroupLabel="Runbooks" ValueLabel="True">
      <Condition Expression="HasCategory('Runbooks')" />
      <Setter Property="Background" Value="#FFC0C0FF" />
    </Style>
    <Style TargetType="Node" GroupLabel="ProductRunbooks" ValueLabel="True">
      <Condition Expression="HasCategory('ProductRunbooks')" />
      <Setter Property="Background" Value="#FFA0A0FF" />
    </Style>
    <Style TargetType="Node" GroupLabel="NestedRunbooks" ValueLabel="True">
      <Condition Expression="HasCategory('NestedRunbooks')" />
      <Setter Property="Background" Value="#FF8080FF" />
    </Style>
  </Styles>
</DirectedGraph>
"@

function New-Graph()
{
    return [xml]$DirectedGraphTemplate
}

function Set-GraphTitle([xml]$xml, [string]$title = $null)
{
    $xml.DirectedGraph.Title = $title
}

function Add-GraphNode([xml]$xml, [string]$id, [string]$label, [string]$category = $null, [string]$group = $null)
{
    # <Node Id="{id}" Label="{label}" Group="(Expanded)" />
    if ($label -eq $null -or $label -eq '') { $label = $id }
    $node = $xml.CreateElement('Node', $Dgml)
    $node.SetAttribute('Id', $id)
    $node.SetAttribute('Label', $label)
    if ($category -ne $null) { $node.SetAttribute('Category', $category) }
    if ($group -ne $null) { $node.SetAttribute('Group', $group) }
    return $xml.DirectedGraph.Nodes.AppendChild($node)
}

function Add-GraphNodeCategory([xml]$xml, [System.Xml.XmlElement]$node, [string]$ref)
{
    # <Category Ref="{category}" />
    $category = $xml.CreateElement('Category', $Dgml)
    $category.SetAttribute('Ref', $ref)
    return $node.AppendChild($category)
}

function Add-GraphLink([xml]$xml, [string]$source, [string]$target, [string]$label = $null, [string]$category = $null)
{
    # <Link Source="{source}" Target="{target}" Category="(Contains|References)" />
    $link = $xml.CreateElement('Link', $Dgml)
    $link.SetAttribute('Source', $source)
    $link.SetAttribute('Target', $target)
    if ($label -ne $null) { $link.SetAttribute('Label', $label) }
    if ($category -ne $null) { $link.SetAttribute('Category', $category) }
    return $xml.DirectedGraph.Links.AppendChild($link)
}

##########################################################################################
# GUI definition
# Events handling buttons and clicks are also added here
########################################################################################## 

# Form and GUI definition
[XML]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        ResizeMode="NoResize"
        Title="SMART Visualization Tool" Height="765" Width="520">
        
                    <Grid>

                        <Grid.RowDefinitions>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="300"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                        </Grid.RowDefinitions>
                        <Label Content="SMA Endpoint :" VerticalAlignment="Center" Grid.Row="0"></Label>
                        <TextBox Text="localhost" Name="DBServer" Grid.Row="0" HorizontalAlignment="Center" VerticalAlignment="Center"  Margin="-100,0,0,0" Width="200"></TextBox>
                        <TextBox Text="9090" Name="DBPort" Grid.Row="0" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="160,0,0,0" Width="35"></TextBox>
                        <Button IsDefault="true" Content="Update List" Name="UpdateWSEndpoint" HorizontalAlignment="Right" VerticalAlignment="Center" Width="100" Margin="0,0,10,0" Grid.Row="0"/>
        
                        <Label Content="1. Pick your Runbook(s) : " FontWeight="Bold" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="7,0,0,0" Grid.Row="1"></Label>
                        <Label Content="You can use this tag-sorted view..." HorizontalAlignment="Left" VerticalAlignment="Center" Margin="160,0,0,0" Grid.Row="1"></Label>
                        <Button Content="Clear selection" Name="ClearTreeSelection" HorizontalAlignment="Right" VerticalAlignment="Center" Width="100" Margin="0,0,10,0" Grid.Row="1"/>
                        <TreeView Name="Tree" Grid.Row="2" Margin="2"/>
                
                        <Label Content="... and/or multi-select some tags" HorizontalAlignment="Left" VerticalAlignment="Center" Width="300" Margin="7,0,0,0" Grid.Row="3"></Label>
                        <ComboBox Name="PickTagsComboBox" HorizontalAlignment="Left" VerticalAlignment="Center" Width="240" Margin="250,0,0,0" Grid.Row="3">
                                <ComboBox.ItemTemplate>
                                    <DataTemplate>
                                        <StackPanel Orientation="Horizontal">
                                            <CheckBox Margin="5" IsChecked="{Binding TagChecked}"/>
                                            <TextBlock Margin="5" Text="{Binding TagName}"/>
                                        </StackPanel>
                                    </DataTemplate>
                            </ComboBox.ItemTemplate>
                        </ComboBox>

                        <Label Content="... and/or select a local folder with PS1 files" HorizontalAlignment="Left" VerticalAlignment="Center" Width="300" Margin="7,0,0,0" Grid.Row="4"></Label>
                        <TextBox Name="PickFolderTextBoxName" IsEnabled="True" Text="" HorizontalAlignment="Left" VerticalAlignment="Center" Width="240" Margin="250,0,0,0" Grid.Row="4"></TextBox>
                        <Label Name="PickFolderLabelNumber" IsEnabled="False" Content="" HorizontalAlignment="Left" VerticalAlignment="Center" Width="140" Margin="250,0,0,0" Grid.Row="5"></Label>
                        <Button Content="Browse..." Name="PickFolderButton" HorizontalAlignment="Right" VerticalAlignment="Center" Width="100" Margin="0,0,10,0" Grid.Row="5"/>

                        <Label Content="2. Choose your options and export with name" FontWeight="Bold" HorizontalAlignment="Left" VerticalAlignment="Center" Width="300" Margin="7,0,0,0" Grid.Row="6"></Label>
                        <TextBox Name="ExportName" IsEnabled="True" Text="SMAExport" HorizontalAlignment="Left" VerticalAlignment="Center" Width="190" Margin="300,0,0,0" Grid.Row="6"></TextBox>
                        <Button Content="Visualize Dependencies" Width="150" IsEnabled="False" Height="150" HorizontalAlignment="Left" Margin="7,0,0,0" Name="ExportButton" Grid.Row="7" Grid.RowSpan="6"/>
                        <CheckBox Content="In PowerShell Console" IsChecked="True" IsEnabled="True" Name="ExportPSConsole" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="200,0,0,0" Grid.Row="7"></CheckBox>
                        <CheckBox Content="In Visio" IsEnabled="False" Name="ExportVSDCB" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="200,0,0,0" Grid.Row="8"></CheckBox>
                        <CheckBox Content="Save and Close" IsEnabled="False" Name="SaveAndCloseVSD" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="215,0,0,0" Grid.Row="9"></CheckBox>
                        <Button Content="Visio Settings ..." Width="150" IsEnabled="False" Height="24" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="330,0,0,0" Name="VisioSettings" Grid.Row="9"/>
                        <CheckBox Content="In Visual Studio" IsEnabled="False" Name="ExportVSCB" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="200,0,0,0" Grid.Row="10"></CheckBox>
                        <CheckBox Content="Save and Close" IsEnabled="False" Name="SaveAndCloseVS" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="215,0,0,0" Grid.Row="11"></CheckBox>
                        <CheckBox Content="In Word" IsEnabled="False" Name="ExportDOCCB" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="200,0,0,0" Grid.Row="12"></CheckBox>
                        <CheckBox Content="Save and Close" IsEnabled="False" Name="SaveAndCloseDOC" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="215,0,0,0" Grid.Row="13"></CheckBox>
                        <Button Content="Visualize All Runbooks" Width="150" IsEnabled="False" Height="27" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="7,0,0,0" Name="ExportAllButton" Grid.Row="13"/>
                        <Button Content="List All Variables" Width="150" IsEnabled="False" Height="27" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="7,0,0,0" Name="ListVariables" Grid.Row="14"/>
                    </Grid>

       
    
</Window>

'@

$Reader = (New-Object System.XML.XMLNodeReader $XAML)
$FORM = [Windows.Markup.XAMLReader]::Load($Reader)
# Linking variables to the GUI
$Global:Tree = New-Object System.Windows.Controls.TreeView
$Global:Tree = $FORM.FindName('Tree')
$Global:Tree.IsEnabled= $False
$Global:SkeletonCB = $FORM.FindName('SkeletonOnly')
$DBServerTB = $FORM.FindName('DBServer')
$DBPortTB = $FORM.FindName('DBPort')
$Global:SaveAndClosePS1 = $FORM.FindName('SaveAndClosePS1')

$FORM.FindName('VisioSettings').Add_Click({
    $ReaderVisioSettings = (New-Object System.XML.XMLNodeReader $XAMLVisioSettings)
    $FORMVisioSettings = [Windows.Markup.XAMLReader]::Load($ReaderVisioSettings)

    $FORMVisioSettings.FindName('VisioTemplateTextBox').Text = $Global:VisioTemplate

    $FORMVisioSettings.FindName('VisioGlueFromComboBox').Items.Add("Bottom") | out-null
    $FORMVisioSettings.FindName('VisioGlueFromComboBox').Items.Add("Top") | out-null
    $FORMVisioSettings.FindName('VisioGlueFromComboBox').Items.Add("Left") | out-null
    $FORMVisioSettings.FindName('VisioGlueFromComboBox').Items.Add("Right") | out-null
    $FORMVisioSettings.FindName('VisioGlueFromComboBox').SelectedIndex = $FORMVisioSettings.FindName('VisioGlueFromComboBox').Items.IndexOf($Global:VisioGlueFrom)

    $FORMVisioSettings.FindName('VisioGlueToComboBox').Items.Add("Bottom") | out-null
    $FORMVisioSettings.FindName('VisioGlueToComboBox').Items.Add("Top") | out-null
    $FORMVisioSettings.FindName('VisioGlueToComboBox').Items.Add("Left") | out-null
    $FORMVisioSettings.FindName('VisioGlueToComboBox').Items.Add("Right") | out-null
    $FORMVisioSettings.FindName('VisioGlueToComboBox').SelectedIndex = $FORMVisioSettings.FindName('VisioGlueToComboBox').Items.IndexOf($Global:VisioGlueTo)


        $FORMVisioSettings.FindName('VisioStencilComboBox').Clear()
        $FORMVisioSettings.FindName('VisioCalloutComboBox').Items.Clear()

        $vApp = New-Object -ComObject Visio.InvisibleApp
        
        $vDoc = $vApp.Documents.Add("")

        $vStencil = $vApp.Documents.OpenEx($FORMVisioSettings.FindName('VisioTemplateTextBox').Text, 4)
        foreach ($vFlowChartMaster in $vStencil.Masters)
            {$FORMVisioSettings.FindName('VisioStencilComboBox').Items.Add($vFlowChartMaster.Name)}
        $vDoc.Close()

        If ($FORMVisioSettings.FindName('VisioStencilComboBox').Items -contains "Process")
            {$FORMVisioSettings.FindName('VisioStencilComboBox').SelectedIndex = $FORMVisioSettings.FindName('VisioStencilComboBox').Items.IndexOf($Global:VisioStencil)}

        $vDoc2 = $vApp.Documents.OpenEx($vApp.GetBuiltInStencilFile(3, 2), 64)
        foreach ($vFlowChartMaster In $vDoc2.Masters)
            {$FORMVisioSettings.FindName('VisioCalloutComboBox').Items.Add($vFlowChartMaster.Name)}
        $vDoc2.Close()
        If ($FORMVisioSettings.FindName('VisioCalloutComboBox').Items -contains "Word Balloon")
            {$FORMVisioSettings.FindName('VisioCalloutComboBox').SelectedIndex = $FORMVisioSettings.FindName('VisioCalloutComboBox').Items.IndexOf($Global:VisioCallout)}

        $vApp.Quit()

        $FORMVisioSettings.ShowDialog() | Out-Null

    $Global:VisioTemplate = $FORMVisioSettings.FindName('VisioTemplateTextBox').Text
    $Global:VisioStencil = $FORMVisioSettings.FindName('VisioStencilComboBox').Text
    $Global:VisioCallout = $FORMVisioSettings.FindName('VisioCalloutComboBox').Text
    $Global:VisioGlueFrom = $FORMVisioSettings.FindName('VisioGlueFromComboBox').Text
    $Global:VisioGlueTo = $FORMVisioSettings.FindName('VisioGlueToComboBox').Text
    $FORMVisioSettings.Close()
    SaveConfig  -ConfigFileLocation $ConfigFileLocation -WSEndPoint $Global:SMAWSEndPoint -WSPort $Global:SMAWSPort -VTemplate $Global:VisioTemplate -VStencil $Global:VisioStencil -VCallout $Global:VisioCallout -VGlueFrom $Global:VisioGlueFrom -VGlueTo $Global:VisioGlueTo
    })
    
$FORM.FindName('ListVariables').Add_Click({
    write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Querying variables..."
    $Variables = get-smavariable -WebServiceEndpoint $Global:SMAWSEndPoint -Port $Global:SMAWSPort | ? name -like "*"
    write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Displaying variables..."
    write-host ($Variables | Select name, isencrypted, value | out-string)
})

$FORM.FindName('ExportAllButton').Add_Click({

    write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Caching resources and determining list of Runbooks to export...(Step 1 of 3)"
    #Get Full List of Runbooks in the system, for performance and caching purposes
    CacheFullRunbookList  

    If ($RunbooksCacheOptimization -eq $False)
    {
    write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Working with" $Global:FullRunbookList.Count "Runbooks"
    ExportRoutine($Global:FullRunbookList)  
    }
    else
    {
    $RunbooksArray = @()
    write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Per the script parameters, extra optimization steps will happen to cache runbooks script..."
    $RunbooksCacheDirectory = $PSCommandPath.Substring(0,$PSCommandPath.LastIndexOf("\")) + "\RunbooksCacheOptimization\"
    If ((Test-Path $RunbooksCacheDirectory) -ne $True) {New-Item $RunbooksCacheDirectory -ItemType Directory}
    $i=1
    foreach ($Runbook in ($Global:FullRunbookList)) 
        {
        Write-Progress -Activity "Runbook Caching Optimization..." -PercentComplete (100*$i/($Global:FullRunbookList.Count))
        (get-smarunbookdefinition -Name $Runbook.RunbookName -WebServiceEndpoint  $Global:SMAWSEndPoint -Port $Global:SMAWSPort -type published).Content| Set-Content ($RunbooksCacheDirectory + $Runbook.RunbookName + ".ps1")
        $RunbooksArray += New-Object PSObject -Property @{ 
             RunbookName = $Runbook.RunbookName
             RunbookID = "LOCALFILE-LRO"
             } 
            $i++
        }
    Write-Progress -Activity "Runbook Caching Optimization..." -Completed
    write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Working with" $RunbooksArray.Count "Runbooks"
    ExportRoutine($RunbooksArray)
    If ($RunbooksCacheOptimization)
        {Remove-Item $RunbooksCacheDirectory -Force -Recurse}
    }      
})

$FORM.FindName('ExportButton').Add_Click({
    
    write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Caching resources and determining list of Runbooks to export...(Step 1 of 3)"
    #Get Full List of Runbooks in the system, for performance and caching purposes
    CacheFullRunbookList

    If ($RunbooksCacheOptimization)
    {
    write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Per the script parameters, extra optimization steps will happen to cache runbooks script..."
    $RunbooksCacheDirectory = $PSCommandPath.Substring(0,$PSCommandPath.LastIndexOf("\")) + "\RunbooksCacheOptimization\"
    If ((Test-Path $RunbooksCacheDirectory) -ne $True) {New-Item $RunbooksCacheDirectory -ItemType Directory}
    }

    $RunbooksArray = @()
    $RunbooksArray0 = @()

     $CurrentNode = $Tree.SelectedItem
    If ($CurrentNode)
        {
        If ($CurrentNode.Name.Contains("Runbook")) {
            #From the tree view, we get a Runbook
            $RunbooksArray0 += New-Object PSObject -Property @{ 
                    RunbookName = $CurrentNode.Header
                    RunbookID = $CurrentNode.Tag
                    }  
            }
        If ($CurrentNode.Name.Contains("Folder")) {
            #From the tree view, we get a 'tag' folder and need to retrieve all the Runbooks it contains
            foreach ($Node in $CurrentNode.Items)
                {
                $RunbooksArray0 += New-Object PSObject -Property @{ 
                    RunbookName = $Node.Header
                    RunbookID = $Node.Tag
                    }  
                }
            }
        }
    #Now we add any Runbooks for which a tag may have been checked in the drop-down list
    $NoTagsRunbooksRequested = $false
    $CheckedTags = ($FORM.FindName('PickTagsComboBox').Items | ? TagChecked -eq $True).TagName
    If ($CheckedTags)
        {
        If ($CheckedTags -contains "[No Tags]") {$NoTagsRunbooksRequested = $true}
        [regex] $CheckedTagsRegEx = ‘(‘ + (($CheckedTags |foreach {[regex]::escape($_)}) –join “|”) + ‘)’
        foreach ($Runbook in ($Global:FullRunbookList | Where-Object {$_.Tags -match $CheckedTagsRegEx})) 
            {
            $RunbooksArray0 += New-Object PSObject -Property @{ 
                 RunbookName = $Runbook.RunbookName
                 RunbookID = $Runbook.RunbookID
                 } 
            }
        }
    #We also  we may have to take into account runbooks without any tags
    If ($NoTagsRunbooksRequested)
        {
        foreach ($Runbook in ($Global:FullRunbookList | Where-Object {$_.Tags -eq $null})) 
            {
            $RunbooksArray0 += New-Object PSObject -Property @{ 
                 RunbookName = $Runbook.RunbookName
                 RunbookID = $Runbook.RunbookID
                 } 
            }
        }
              
    #If $RunbooksCacheOptimization is set to $True, we export their script content and "flag" runbooks as such
    If ($RunbooksCacheOptimization)
    {
    for ($i=0;$i -lt $RunbooksArray0.Count;$i++)
        {
        Write-Progress -Activity "Runbook Caching Optimization..." -PercentComplete (100*$i/($RunbooksArray0.Count))
        (get-smarunbookdefinition -Name $RunbooksArray0[$i].RunbookName -WebServiceEndpoint  $Global:SMAWSEndPoint -Port $Global:SMAWSPort -type published).Content| Set-Content ($RunbooksCacheDirectory + $RunbooksArray0[$i].RunbookName + ".ps1")
        $RunbooksArray0[$i].RunbookID = "LOCALFILE-LRO"
        }
    Write-Progress -Activity "Runbook Caching Optimization..." -Completed
    }

    #Finally, a folder may have been selected too
    If ($FORM.FindName('PickFolderTextBoxName').Text -ne "")
        {
        foreach ($RunbookFile in (Get-Item ($FORM.FindName('PickFolderTextBoxName').Text + "\*.ps1"))) 
            {
            #This assumes runbooks file names are the same as the workflow names, and that they are indeed workflows (not just standard PS1)
            $RunbooksArray0 += New-Object PSObject -Property @{ 
                 RunbookName = $RunbookFile.Name.Substring(0,$RunbookFile.Name.Length-4)
                 RunbookID = "LOCALFILE"
                 } 
            }
        }  

    #Selecting unique Runbooks names
    $RunbooksArray = @()
    foreach ($Item in $RunbooksArray0){If (($RunbooksArray.RunbookName -contains $Item.RunbookName) -eq $false){$RunbooksArray+= $Item}}

    #Run Export Routine
    If ($RunbooksArray.Count -eq 0)
        {
        write-host -ForegroundColor red "["(date -format "HH:mm:ss")"] WARNING : No Runbooks were selected in the tool."
        popup2 -Message "No Runbooks were selected in the tool." -ClosedExternally $False -NbLines 2 -ShowDialog $True
        }
        else
        {
        write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Working with" $RunbooksArray.Count "Runbooks"
        ExportRoutine($RunbooksArray)
        }
    If ($RunbooksCacheOptimization)
        {Remove-Item $RunbooksCacheDirectory -Force -Recurse}
})

$FORM.FindName('UpdateWSEndpoint').Add_Click({
        $Global:SMAWSEndPoint = $FORM.FindName('DBServer').Text
        $Global:SMAWSPort = $FORM.FindName('DBPort').Text
        ListRunbooks($Global:Tree)
})

$FORM.FindName('PickFolderButton').Add_Click({

    
    $NewShell = New-Object -comObject Shell.Application   
    $PickedFolder = $NewShell.BrowseForFolder(0, "Pick a folder with PS1 files", 0, 0)  
    if ($PickedFolder -ne $null) {  
        $PS1Folder = $PickedFolder.self.Path
        #$FORM.FindName('PickFolderLabel').Content = $PS1Folder
        $NumberOfRunbooksInPickedfolder = (Get-Item "$PS1Folder\*.ps1").Count
        $FORM.FindName('PickFolderTextBoxName').Text = $PS1Folder
        $FORM.FindName('PickFolderLabelNumber').Content = "[$NumberOfRunbooksInPickedfolder Runbook(s) found]"
    }  
    else
    {
    write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] No Folder selected by user, returning to main window..."
    }
})

$FORM.FindName('ClearTreeSelection').Add_Click({

       $Count = 0
       Write-Progress -Activity "Clearing list..." -PercentComplete $Count
       If ($Count -gt 100) { $Count = 0}
       foreach ($TreeNode in $Tree.Items)
              {
              $Count = $Count + 10
              $TreeNode.IsSelected = $false
              ClearNode($TreeNode)
              }
       Write-Progress -Activity "Clearing list..." -Completed
})

#Only enable Visio, Word and Visual Studio options is they are installed

If (Test-Path "HKLM:\Software\Classes\.doc\Word.Document.8\")
    {
    $FORM.FindName('ExportDOCCB').IsEnabled = $True
    #$FORM.FindName('SingleWordDoc').IsEnabled = $True
    $FORM.FindName('SaveAndCloseDOC').IsEnabled = $True
    }

If (Test-Path "HKLM:\Software\Classes\.vsd\Visio.Drawing.11\")
    {
    $FORM.FindName('ExportVSDCB').IsEnabled = $True
    $FORM.FindName('VisioSettings').IsEnabled = $True
    $FORM.FindName('SaveAndCloseVSD').IsEnabled = $True
    #$FORM.FindName('GroupThumbnails').IsEnabled = $True
    }

 If (Test-Path "HKLM:\Software\Classes\.dgml\")
    {
    $FORM.FindName('ExportVSCB').IsEnabled = $True
    $FORM.FindName('SaveAndCloseVS').IsEnabled = $True
    }


##########################################################################################
# Actual start of the main process
##########################################################################################

cls
write-host -ForegroundColor green "["(date -format "HH:mm:ss")"] SMART Visualization Tool $ToolVersion"
# We default to the web service endpoint and port values from the script, as well as initialize Visio variables
$Global:SMAWSEndPoint = $WSEndPoint
$Global:SMAWSPort = $WSPort
$Global:VisioTemplate = "C:\Program Files (x86)\Microsoft Office\Office15\Visio Content\1033\BASFLO_U.VSSX"
$Global:VisioStencil = "Process"
$Global:VisioCallout = "Word Balloon"
$Global:VisioGlueFrom = "Bottom"
$Global:VisioGlueTo = "Top"
# Let's try to see if there is a config file and, if yes, update the global variables
$ConfigFileLocation = (Get-Location -PSProvider FileSystem).Path + "\Config-SMART-VT.xml"
If (Test-Path -Path $ConfigFileLocation){LoadConfig -ConfigFileLocation $ConfigFileLocation}
$DBServerTB.Text = $Global:SMAWSEndPoint
$DBPortTB.Text = $Global:SMAWSPort
#Fill out the GUI and display the form
ListRunbooks($Global:Tree)
write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Displaying GUI..."
$FORM.ShowDialog() | Out-Null
# Everything after  this line is executed when exiting the tool
SaveConfig  -ConfigFileLocation $ConfigFileLocation -WSEndPoint $Global:SMAWSEndPoint -WSPort $Global:SMAWSPort -VTemplate $Global:VisioTemplate -VStencil $Global:VisioStencil -VCallout $Global:VisioCallout -VGlueFrom $Global:VisioGlueFrom -VGlueTo $Global:VisioGlueTo
write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Exiting..."

