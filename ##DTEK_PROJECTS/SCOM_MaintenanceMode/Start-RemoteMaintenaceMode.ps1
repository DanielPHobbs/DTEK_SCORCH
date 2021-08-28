#Generated Form Function
function GenerateForm {
########################################################################
# Code Generated By: SAPIEN Technologies PrimalForms (Community Edition) v1.0.10.0
# Generated On: 07.10.2012 22:34
# Generated By: roth_000
########################################################################

#region Import the Assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
#endregion

######################################################################
#Define Orchestrator webservice server
$Webserver="dtekorch16-s1"
#Define the Runbook which you would like to run => GUID
[GUID]$RunbookGUID="85c43eef-d142-47de-9935-3cb3346fa947"
#Define Runbook server
$RunbookServer="dtekorch16-s1"
#Path to the SCOJobRunner tool
$Path="C:\SCORCH"
######################################################################

#region Generated Form Objects
$frmMM = New-Object System.Windows.Forms.Form
$label1 = New-Object System.Windows.Forms.Label
$btnCancel = New-Object System.Windows.Forms.Button
$lblServerSum = New-Object System.Windows.Forms.Label
$lblSummary = New-Object System.Windows.Forms.Label
$btnOK = New-Object System.Windows.Forms.Button
$MMGroup = New-Object System.Windows.Forms.GroupBox
$cmbOS = New-Object System.Windows.Forms.ComboBox
$lblTime = New-Object System.Windows.Forms.Label
$cmbTime = New-Object System.Windows.Forms.ComboBox
$lblReason = New-Object System.Windows.Forms.Label
$cmbReason = New-Object System.Windows.Forms.ComboBox
$lblComment = New-Object System.Windows.Forms.Label
$txtComment = New-Object System.Windows.Forms.TextBox
$lblServer = New-Object System.Windows.Forms.Label
$txtServer = New-Object System.Windows.Forms.TextBox
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#Load values into the comboxes (cmbTime=Minutes, cmbReason=reason for MM)
$cmbTime.Items.Add(10)
$cmbTime.Items.Add(20)
$cmbTime.Items.Add(30)
$cmbTime.Items.Add(40)
$cmbReason.Items.Add("PlannedOther")
$cmbReason.Items.Add("UnplannedOther")
$cmbReason.Items.Add("SecurityIssue")
$cmbOS.Items.Add("Windows Maintenance Mode and Restart")
$cmbOS.Items.Add("Windows Maintenance Mode Only")
$cmbOS.Items.Add("Linux Maintenance Mode Only")

#Action when OK button is pressed
$btnOK_OnClick= 
{
#Get value from cmbReason
$Reason = $cmbReason.SelectedItem.ToString()

#Get value from cmbTime
$Time = $cmbTime.SelectedItem

#Get comment from txtcomment field, if empty set current logged in user
$Comment = If ($txtComment.Text.Length -eq 0) {$txtComment.Text = "Server boot by $env:Username"}

#Check if server text field is not empty, if true write output
If ($txtServer.Text.Length -eq 0 ) 

    {
    Write-Host -ForegroundColor Red "Please enter a server name FQDN!"
    }
       
    Else 
    
    {
    #Show Summary label and text
    $lblSummary.Show()
    $lblServerSum.Show()
    $lblServerSum.Text=$txtServer.Text
    $Object= $txtServer.Text
    $OS=$cmbOS.SelectedItem.ToString()

    #Call the SCOJobRunner Tool with parameters
    Invoke-Expression -command "&`"$Path\SCOJobRunner.exe`" -ID:$RunbookGUID -Webserver:$Webserver -RunbookServer:$RunbookServer '-Parameters:Object=$Object;Time=$Time;Reason=$Reason;Comment=$Comment;User=$env:Username;OS=$OS'"
    }

}

$handler_groupBox1_Enter= 
{
#TODO: Place custom script here

}

$btnCancel_OnClick= 
{
#TODO: Place custom script here
$frmMM.Close()

}

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$frmMM.WindowState = $InitialFormWindowState
}


#----------------------------------------------
#region Generated Form Code
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 447
$System_Drawing_Size.Width = 282
$frmMM.ClientSize = $System_Drawing_Size
$frmMM.DataBindings.DefaultDataSourceUpdateMode = 0
$frmMM.Name = "frmMM"
$frmMM.Text = "SCOM Maintenance Mode"

$label1.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 173
$System_Drawing_Point.Y = 429
$label1.Location = $System_Drawing_Point
$label1.Name = "label1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 99
$label1.Size = $System_Drawing_Size
$label1.TabIndex = 5
$label1.Text = "(C) by SCOMfaq"
$label1.TextAlign = 4
$label1.add_Click($handler_label1_Click)

$frmMM.Controls.Add($label1)


$btnCancel.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 159
$System_Drawing_Point.Y = 400
$btnCancel.Location = $System_Drawing_Point
$btnCancel.Name = "btnCancel"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$btnCancel.Size = $System_Drawing_Size
$btnCancel.TabIndex = 4
$btnCancel.Text = "Cancel"
$btnCancel.UseVisualStyleBackColor = $True
$btnCancel.add_Click($btnCancel_OnClick)

$frmMM.Controls.Add($btnCancel)

$lblServerSum.BackColor = [System.Drawing.Color]::FromArgb(255,185,209,234)
$lblServerSum.DataBindings.DefaultDataSourceUpdateMode = 0
$lblServerSum.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,1,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 19
$System_Drawing_Point.Y = 358
$lblServerSum.Location = $System_Drawing_Point
$lblServerSum.Name = "lblServerSum"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 240
$lblServerSum.Size = $System_Drawing_Size
$lblServerSum.TabIndex = 3
$lblServerSum.Visible = $False

$frmMM.Controls.Add($lblServerSum)

$lblSummary.BackColor = [System.Drawing.Color]::FromArgb(255,185,209,234)
$lblSummary.DataBindings.DefaultDataSourceUpdateMode = 0
$lblSummary.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,1,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 20
$System_Drawing_Point.Y = 317
$lblSummary.Location = $System_Drawing_Point
$lblSummary.Name = "lblSummary"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 22
$System_Drawing_Size.Width = 240
$lblSummary.Size = $System_Drawing_Size
$lblSummary.TabIndex = 2
$lblSummary.Text = "Server selected..."
$lblSummary.Visible = $False

$frmMM.Controls.Add($lblSummary)


$btnOK.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 41
$System_Drawing_Point.Y = 400
$btnOK.Location = $System_Drawing_Point
$btnOK.Name = "btnOK"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$btnOK.Size = $System_Drawing_Size
$btnOK.TabIndex = 1
$btnOK.Text = "OK"
$btnOK.UseVisualStyleBackColor = $True
$btnOK.Visible = $True
$btnOK.add_Click($btnOK_OnClick)

$frmMM.Controls.Add($btnOK)


$MMGroup.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 13
$MMGroup.Location = $System_Drawing_Point
$MMGroup.Name = "MMGroup"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 288
$System_Drawing_Size.Width = 259
$MMGroup.Size = $System_Drawing_Size
$MMGroup.TabIndex = 0
$MMGroup.TabStop = $False
$MMGroup.Text = "Please select..."
$MMGroup.add_Enter($handler_groupBox1_Enter)

$frmMM.Controls.Add($MMGroup)
$cmbOS.DataBindings.DefaultDataSourceUpdateMode = 0
$cmbOS.FormattingEnabled = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 241
$cmbOS.Location = $System_Drawing_Point
$cmbOS.Name = "cmbOS"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 246
$cmbOS.Size = $System_Drawing_Size
$cmbOS.TabIndex = 8
$cmbOS.Text = "Windows Maintenance Mode Only"

$MMGroup.Controls.Add($cmbOS)

$lblTime.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 77
$lblTime.Location = $System_Drawing_Point
$lblTime.Name = "lblTime"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 19
$System_Drawing_Size.Width = 248
$lblTime.Size = $System_Drawing_Size
$lblTime.TabIndex = 7
$lblTime.Text = "Duration (Min)"

$MMGroup.Controls.Add($lblTime)

$cmbTime.DataBindings.DefaultDataSourceUpdateMode = 0
$cmbTime.FormattingEnabled = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 5
$System_Drawing_Point.Y = 96
$cmbTime.Location = $System_Drawing_Point
$cmbTime.Name = "cmbTime"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 121
$cmbTime.Size = $System_Drawing_Size
$cmbTime.TabIndex = 1
$cmbTime.Text = "10"

$MMGroup.Controls.Add($cmbTime)

$lblReason.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 126
$lblReason.Location = $System_Drawing_Point
$lblReason.Name = "lblReason"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 244
$lblReason.Size = $System_Drawing_Size
$lblReason.TabIndex = 5
$lblReason.Text = "Reason"

$MMGroup.Controls.Add($lblReason)

$cmbReason.DataBindings.DefaultDataSourceUpdateMode = 0
$cmbReason.FormattingEnabled = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 148
$cmbReason.Location = $System_Drawing_Point
$cmbReason.Name = "cmbReason"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 247
$cmbReason.Size = $System_Drawing_Size
$cmbReason.TabIndex = 3
$cmbReason.Text = "SecurityIssue"

$MMGroup.Controls.Add($cmbReason)

$lblComment.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 5
$System_Drawing_Point.Y = 182
$lblComment.Location = $System_Drawing_Point
$lblComment.Name = "lblComment"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 19
$System_Drawing_Size.Width = 243
$lblComment.Size = $System_Drawing_Size
$lblComment.TabIndex = 3
$lblComment.Text = "Comment"

$MMGroup.Controls.Add($lblComment)

$txtComment.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 202
$txtComment.Location = $System_Drawing_Point
$txtComment.Name = "txtComment"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 246
$txtComment.Size = $System_Drawing_Size
$txtComment.TabIndex = 4

$MMGroup.Controls.Add($txtComment)

$lblServer.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 28
$lblServer.Location = $System_Drawing_Point
$lblServer.Name = "lblServer"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 244
$lblServer.Size = $System_Drawing_Size
$lblServer.TabIndex = 1
$lblServer.Text = "FQDN Server"

$MMGroup.Controls.Add($lblServer)

$txtServer.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 46
$txtServer.Location = $System_Drawing_Point
$txtServer.Name = "txtServer"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 247
$txtServer.Size = $System_Drawing_Size
$txtServer.TabIndex = 0

$MMGroup.Controls.Add($txtServer)


#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $frmMM.WindowState
#Init the OnLoad event to correct the initial state of the form
$frmMM.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$frmMM.ShowDialog()| Out-Null

} #End Function

#Call the Function
GenerateForm
