#######################################################################
# PowerShell WinForm Designer
# Requires PowerShell 2.0 or later
######################################################################
#
# Changelog:
# 1.0     - 4/24/2015 Began as "PowerShell Form Builder" by Z.Alex <https://gallery.technet.microsoft.com/scriptcenter/Powershell-Form-Builder-3bcaf2c7>
# 2.1.1   - Codebase is based on the Version 2.1.1 by mozers <https://bitbucket.org/mozers/powershell-formdesigner>
# ------
# TODO:
# 1. Add the ability to choose between saving the form for use as standalone form code (for use in .dotsourcing the form into a main script file) or
#    as a complete codeset with form code and main script code in one.
# 2. Refactor Properties and events section to streamline their functions and variable use
# 3. Redesign Main form to allow for addition of adding more events to events section along with code specific to that event
# 4. Add ability for tabbed forms and multi window forms
# 5. Set message when property is left blank
# 6. Clear Main Form on Design form close... (Clear out datagridviews)
#-------
# Notes: Anytime a control is added to the form or code is used to add or change properties, it is really just adding it to the Control Grid list box in the Controls Section.
# It does show up on the new form, but that is just a physical representation of the controls listed in that box. The code will perform functions on and sets and changes
# things to the items in that box, not the form itself.
#
# Control Variable names: Prefix is type of control e.g. -> dgv is Data Grid View, cb is Combo Box, etc.
# GUI Element Variables name: Prefix is type of element e.g -> b is Button, cb Combo Box, gb is Group Box, lab is Label, etc.
#
#region === Initialization ===

#requires -version 2.0

#--- Load Assemblies ---

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore, PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic
[Windows.Forms.Application]::EnableVisualStyles()

#--- Global Variables ---

$Global:frmDesign = $null
$Global:currentCtrl = $null
$Global:source = ''
[Int32]$Global:curFirstX = 0
[Int32]$Global:curFirstY = 0
[string]$Version = '2.5'

#endregion === Initialization ===

#region === Functions ===

#--- Global Functions ---
function Open-FilenameDialog($dlgName) {
     # Configures Save and Open Form Dialog box
     $dialog = New-Object System.Windows.Forms.$dlgName
     $dialog.Filter = "Powershell Script (*.ps1)|*.ps1|All files (*.*)|*.*"
     $dialog.ShowDialog() | Out-Null
     return $dialog.filename
}

#--- Form Functions ---
function Set-defaultsPSFDMain {
     $bSaveForm.Enabled = $false
     $bAddControl.Enabled = $false
     $bRemoveControl.Enabled = $false
     $dgvControls.Rows.Clear()
     $cbAddControl.SelectedItem = $null
     $bAddProp.Enabled = $false
     $bRemoveProp.Enabled = $false
     $dgvProps.Rows.Clear()
     $cbAddProp.SelectedItem = $null
     $dgvEvents.Rows.Clear()
     $cbAddEvent.SelectedItem = $null
     $bNewForm.Enabled = $true
     $bOpenForm.Enabled = $true
}
function Set-BtnsFrmDesignOpen {
     # New Form Settings
     $bSaveForm.Enabled = $true
     $bAddControl.Enabled = $true
     $bRemoveControl.Enabled = $true
     $bAddProp.Enabled = $true
     $bRemoveProp.Enabled = $true
     $bNewForm.Enabled = $false
     $bOpenForm.Enabled = $false
}
function Close-FrmDesign {
     $ButtonType = [System.Windows.MessageBoxButton]::YesNo
     $MessageboxTitle = "Save Form Design?"
     $Messageboxbody = "Do you want to save this new form design?"
     $MessageIcon = [System.Windows.MessageBoxImage]::Question
     $result = [System.Windows.MessageBox]::Show($Messageboxbody, $MessageboxTitle, $ButtonType, $messageicon)
     switch ($result) {
          'Yes' {
               Save-FrmDesign
               Set-defaultsPSFDMain
          }
          'No' {
               Set-defaultsPSFDMain
          }
     }

}

#--- Resize and Move Control Functions ---
function Set-MouseDown {
     # Sets position on mouse button down
     $Global:curFirstX = ([System.Windows.Forms.Cursor]::Position.X)
     $Global:curFirstY = ([System.Windows.Forms.Cursor]::Position.Y)
}
function Set-MouseMove ($controlName) {
     # Moves the Control along with the movement of the mouse, and monitors its current location
     $curPosX = ([System.Windows.Forms.Cursor]::Position.X)
     $curPosY = ([System.Windows.Forms.Cursor]::Position.Y)
     $borderWidth = ($Global:frmDesign.Width - $Global:frmDesign.ClientSize.Width) / 2
     $titlebarHeight = $Global:frmDesign.Height - $Global:frmDesign.ClientSize.Height - 2 * $borderWidth

     if ($Global:curFirstX -eq 0 -and $Global:curFirstY -eq 0) {
          if ($this.Parent -ne $Global:frmDesign) {
               $groupBoxLocationX = $this.Parent.Location.X
               $groupBoxLocationY = $this.Parent.Location.Y
          }
          else {
               $groupBoxLocationX = 0
               $groupBoxLocationY = 0
          }
          # Sets icon arrows when resizing control
          $isWidthChange = ($curPosX - $Global:frmDesign.Location.X - $groupBoxLocationX - $this.Location.X) -ge $this.Width
          $isHeightChange = ($curPosY - $Global:frmDesign.Location.Y - $groupBoxLocationY - $this.Location.Y) -ge ($this.Height + $titlebarHeight)
          if ($isWidthChange -and $isHeightChange) {
               $this.Cursor = "SizeNWSE"
          }
          elseif ($isWidthChange) {
               $this.Cursor = "SizeWE"
          }
          elseif ($isHeightChange) {
               $this.Cursor = "SizeNS"
          }
          else {
               $this.Cursor = "SizeAll"
          }
     }
     else {
          $difX = $Global:curFirstX - $curPosX
          $difY = $Global:curFirstY - $curPosY
          switch ($this.Cursor) {
               "[Cursor: SizeWE]" {
                    $this.Width = $this.Width - $difX
               }
               "[Cursor: SizeNS]" {
                    $this.Height = $this.Height - $difY
               }
               "[Cursor: SizeNWSE]" {
                    $this.Width = $this.Width - $difX
                    $this.Height = $this.Height - $difY
               }
               "[Cursor: SizeAll]" {
                    $this.Left = $this.Left - $difX
                    $this.Top = $this.Top - $difY
               }
          }
          $Global:curFirstX = $curPosX
          $Global:curFirstY = $curPosY
     }
}
function Set-MouseUP {
     # Sets the cursor to standard icon arrow and sets location on mouse release
     $this.Cursor = "SizeAll"
     $Global:curFirstX = 0
     $Global:curFirstY = 0
     Update-CurrentCrtlPropsGrid
}
function Set-ResizeAndMoveWithKeyboard {
     # Allows resizing and fine movement of current control with keyboard arrow keys
     if ($Global:currentCtrl) {
          if ($_.KeyCode -eq 'Left' -and $_.Modifiers -eq 'None') {
               $Global:currentCtrl.Left -= 1
               $_.Handled = $true
               Update-CurrentCrtlPropsGrid
          }
          elseif ($_.KeyCode -eq 'Left' -and $_.Modifiers -eq 'Control') {
               $Global:currentCtrl.Width -= 1
               $_.Handled = $true
               Update-CurrentCrtlPropsGrid
          }
          elseif ($_.KeyCode -eq 'Right' -and $_.Modifiers -eq 'None') {
               $Global:currentCtrl.Left += 1
               $_.Handled = $true
               Update-CurrentCrtlPropsGrid
          }
          elseif ($_.KeyCode -eq 'Right' -and $_.Modifiers -eq 'Control') {
               $Global:currentCtrl.Width += 1
               $_.Handled = $true
               Update-CurrentCrtlPropsGrid
          }
          elseif ($_.KeyCode -eq 'Up' -and $_.Modifiers -eq 'None') {
               $Global:currentCtrl.Top -= 1
               $_.Handled = $true
               Update-CurrentCrtlPropsGrid
          }
          elseif ($_.KeyCode -eq 'Up' -and $_.Modifiers -eq 'Control') {
               $Global:currentCtrl.Height -= 1
               $_.Handled = $true
               Update-CurrentCrtlPropsGrid
          }
          elseif ($_.KeyCode -eq 'Down' -and $_.Modifiers -eq 'None') {
               $Global:currentCtrl.Top += 1
               $_.Handled = $true
               Update-CurrentCrtlPropsGrid
          }
          elseif ($_.KeyCode -eq 'Down' -and $_.Modifiers -eq 'Control') {
               $Global:currentCtrl.Height += 1
               $_.Handled = $true
               Update-CurrentCrtlPropsGrid
          }
          elseif ($_.KeyCode -eq 'Delete' -and $_.Modifiers -eq 'None') {
               $_.Handled = $true
               Remove-CurrentCtrl
          }
     }
}

#--- Control Functions ---
function Get-CtrlType($ctrl) {
     # Gets the TypeName of the selected control
     return $ctrl.GetType().FullName -replace "System.Windows.Forms.", ""
}
function Remove-CurrentCtrl {
     # Removes a selected Control from the form (From the Grid Box or the form)
     $Global:frmDesign.Controls.Remove($Global:currentCtrl)
     $Global:currentCtrl = $Global:frmDesign
     Update-ControlList
}
function Select-CurrentCtrlInControlList {
     # Selects the newly created control in the Controls list grid box
     $dgvControls.Rows | ForEach-Object {
          if ($_.Cells[0].Value -eq $Global:currentCtrl.Name) {
               $_.Selected = $true
               return
          }
     }
}
function Update-ControlList {
     # Updates view in Controls Grid View Box
     function Get-ContainerControl($container) {
          foreach ($ctrl in $container.Controls) {
               $type = Get-CtrlType $ctrl
               $parent = $ctrl.Parent.Text
               if ($type -eq 'GroupBox') {
                    Get-ContainerControl $ctrl
               }
               $dgvControls.Rows.Add($ctrl.Name, $type, $parent, $ctrl)
          }
     }

     $dgvControls.Rows.Clear()
     $dgvControls.Rows.Add($Global:frmDesign.Name, 'Form', 'TopLevel', $Global:frmDesign)
     Get-ContainerControl $Global:frmDesign
     Select-CurrentCtrlInControlList
}
function Set-CurrentCtrl($Arg) {
     # Sets the selected control on the form (into the Control list grid box)
     if ($arg.GetType().FullName -eq 'System.Int32') {
          $Global:currentCtrl = $dgvControls.Rows[$arg].Cells[3].Value
     }
     else {
          $Global:currentCtrl = $arg
          Select-CurrentCtrlInControlList
     }
     $Global:currentCtrl.Focus()
     Update-AvailablePropsInCB
     Update-CurrentCrtlPropsGrid
     Update-AvailableEventsCB
     Update-CurrentCtrlEventsGrid
}
function Add-Control {
     # Actions to take when the add button in the controls section is clicked, adds control to design and allows configuration
     function Set-CtrlName($ctrlType) {
          # Builds New Control Name and adds a number to end of the name to differentiate Controls (ex. ListBox0, ListBox1, etc.)
          $ctrlNames = $dgvControls.Rows | ForEach-Object {$_.Cells[0].Value}
          $num = 0
          do {
               $newCtrlName = $ctrlType + $num
               $num += 1
          } while ($ctrlNames -contains $newCtrlName)
          return $newCtrlName
     }

     $ctrlType = $cbAddControl.Items[$cbAddControl.SelectedIndex]
     $Control = New-Object System.Windows.Forms.$ctrlType
     $Control.Name = Set-CtrlName $ctrlType
     $Control.Cursor = 'SizeAll'
     if ($ctrlType -eq 'ComboBox') {
          # Disables integralHeight on Combo and list box controls
          $Control.IntegralHeight = $false
     }
     elseif ($ctrlType -eq 'ListBox') {
          $Control.IntegralHeight = $false
     }
     $Control.Tag = @('Name', 'Left', 'Top', 'Width', 'Height') # Adds all standard properties to all controls
     if (@('Button', 'CheckBox', 'GroupBox', 'Label', 'RadioButton') -contains $ctrlType) {
          # Adds special properties for these specific controls
          $Control.Text = $Control.Name
          $Control.Tag += 'Text'
     }
     # Sets the abilities to resize and move the new control with the mouse on the design form (see Functions above)
     $Control.Add_PreviewKeyDown( {$_.IsInputKey = $true} )
     $Control.Add_KeyDown( {Set-ResizeAndMoveWithKeyboard} )
     $Control.Add_MouseDown( {Set-MouseDown} )
     $Control.Add_MouseMove( {Set-MouseMove} )
     $Control.Add_MouseUP( {Set-MouseUP} )
     $Control.Add_Click( {Set-CurrentCtrl $this} )

     $curCtrlType = Get-CtrlType $Global:currentCtrl
     # If newly created control is added to a group area box (Outlined area), code attaches it to that GroupBox as its parent instead of the form itself
     if ($curCtrlType -eq 'GroupBox') {
          $Global:currentCtrl.Controls.Add($Control)
     }
     else {
          $Global:frmDesign.Controls.Add($Control)
          Set-CurrentCtrl $Control
     }

     Update-ControlList
}

#--- Properties Functions ---
function Update-AvailablePropsInCB {
     # Fills in the All Properties Dropdown Box with all available properties of the selected control
     $cbAddProp.Items.Clear()
     $Global:currentCtrl | Get-Member -membertype properties | ForEach-Object {$cbAddProp.Items.Add($_.Name)}
}
function Update-CurrentCrtlPropsGrid {
     # Fills in Current Controls Properties Grid box with all configured properties of the current selected control
     try {
          $dgvProps.Rows.Clear()
     }
     catch {
          return
     }
     [array]$props = $Global:currentCtrl | Get-Member -membertype properties
     foreach ($prop in $props) {
          $propName = $prop.Name
          if ($Global:currentCtrl.Tag -contains $propName) {
               $value = $Global:currentCtrl.$propName
               if ($null -ne $value) {
                    if ($value.GetType().FullName -eq 'System.Drawing.Font') {
                         # Grabs the System font info and adds it to the System.Drawing.Font Properties if needed
                         $value = $value.Name + ', ' + $value.SizeInPoints + ', ' + $value.Style
                    }
               }
               $value = $value -replace 'Color \[(\w+)\]', '$1'
               $dgvProps.Rows.Add($propName, $value)
          }
     }
}
function Add-Property($propName) {
     # Adds Selected property to the Current controls properties Grid box
     $Global:currentCtrl.Tag += $propName
     Update-CurrentCrtlPropsGrid
}
function Remove-Property($propName) {
     # Adds Selected property to the Current control Set properties Grid box
     $Global:currentCtrl.Tag = $Global:currentCtrl.Tag | Where-Object {$_ -ne $propName}
     Update-CurrentCrtlPropsGrid
}
function Set-PropGridCellEndEdit {
     # Fixes Entries in the properties value cell if they are manually edited to ensure proper functionality
     $propName = $dgvProps.CurrentRow.Cells[0].Value
     $value = $dgvProps.CurrentRow.Cells[1].Value

     $allMatches = [regex]::matches($value, '^([\w ]+),\s*(\d+),\s*(\w+)$') # Sets regex to match to find font Values entered.
     if ($allMatches.Success) {
          foreach ($match in $allMatches) {
               # Sets font property
               $fontName = [string]$match.Groups[1]
               $fontSize = [string]$match.Groups[2]
               $fontStyle = [string]$match.Groups[3]
               $Global:currentCtrl.Font = New-Object System.Drawing.Font($fontName, $fontSize, [System.Drawing.FontStyle]::$fontStyle)
          }
     }
     else {
          if ($value -eq 'True') {
               $value = $true
          }
          elseif ($value -eq 'False') {
               $value = $false
          }
          $Global:currentCtrl.$propName = $value
     }
     if ($Global:currentCtrl.Tag -notcontains $propName) {
          $Global:currentCtrl.Tag += $propName
     }
     if ($propName -eq 'Name') {
          Update-ControlList
     }
     Update-CurrentCrtlPropsGrid
}

#--- Event Functions ---
function Update-AvailableEventsCB {
     # Lists all available events to the selected control in the Event Handlers Dropdown box
     $cbAddEvent.Items.Clear()
     $Global:currentCtrl | Get-Member | ForEach-Object {if ($_ -like '*EventHandler*') {
               $cbAddEvent.Items.Add($_.Name)
          }}
}
function Update-CurrentCtrlEventsGrid {
     # Refresh the events Grid box and Adds the Selected event handler of the current control to the list
     $dgvEvents.Rows.Clear()
     [array]$events = $Global:currentCtrl | Get-Member | Where-Object {$_ -like '*EventHandler*'}
     foreach ($eventItem in $events) {
          $eventName = $eventItem.Name
          if ($Global:currentCtrl.Tag -like "Add_$eventName(*") {
               $dgvEvents.Rows.Add($eventName)
          }
     }
}
function Add-Event {
     # On button click, grabs the event handler selected and adds to the current control Tag property.
     $eventH = $cbAddEvent.Items[$cbAddEvent.SelectedIndex]
     $Global:currentCtrl.Tag += 'Add_' + $eventH + '($' + $Global:currentCtrl.Name + '_' + $eventH + ')'
     Update-CurrentCtrlEventsGrid
}

#endregion === Functions ===

#region === Forms ===

#--- New Form ---
function Create-NewFrmDesign {
     # On New form button Click, Creates new design form and sets Form Designer standard settings
     $formName = [Microsoft.VisualBasic.Interaction]::InputBox('Enter a Name for the New Form', 'New Design Form Name?', 'Form0')
     $Global:frmDesign = New-Object System.Windows.Forms.Form
     $Global:frmDesign.Name = "$formName"
     $Global:frmDesign.Text = "$formName"
     $Global:frmDesign.Width = 500
     $Global:frmDesign.Height = 500
     $Global:frmDesign.Tag = @('Name', 'Width', 'Height', 'Text')
     $Global:frmDesign.Add_ResizeEnd( {Update-CurrentCrtlPropsGrid} )
     $Global:frmDesign.Add_FormClosing( {Close-FrmDesign} )
     $Global:frmDesign.StartPosition = "CenterScreen"
     $Global:frmDesign.Show()
     $Global:currentCtrl = $Global:frmDesign
     Update-ControlList
     Update-AvailablePropsInCB
     Update-CurrentCrtlPropsGrid
     Update-AvailableEventsCB
     Update-CurrentCtrlEventsGrid
     Set-BtnsFrmDesignOpen
}

#--- Save Form ---
function Save-FrmDesign {
     # On Save button Click, Configures file to be saved by building all code to be added to the ps1 file
     function Enumerate-SaveControls ($container) {
          # Loop Through Controls to Build Source Code
          [string]$newline = "`r`n"
          [int]$left = 0
          [int]$top = 0
          [int]$width = 0
          [int]$height = 0
          $Global:source += $newline + '#' + $container.Name + $newline # Control code Description/Headline Comment
          [string]$ctrlType = Get-CtrlType $container # Grabs the Controls Type
          $Global:source += '$' + $container.Name + ' = New-Object System.Windows.Forms.' + $ctrlType + $newline
          [array]$props = $container | Get-Member -membertype properties # Get all properties of Control
          foreach ($prop in $props) {
               $propName = $prop.Name
               if ($container.Tag -contains $propName -and $propName -ne "Name") {
                    if ($propName -eq "Left") {
                         $left = $container.Left
                    }
                    elseif ($propName -eq "Top") {
                         $top = $container.Top
                    }
                    elseif ($propName -eq "Width") {
                         $width = $container.Width
                    }
                    elseif ($propName -eq "Height") {
                         $height = $container.Height
                    }
                    else {
                         $value = $container.$propName
                         if ($value.GetType().FullName -eq 'System.Drawing.Font') {
                              $fontName = $value.Name
                              $fontSize = $value.SizeInPoints
                              $fontStyle = $value.Style
                              $value = 'New-Object System.Drawing.Font("' + $fontName + '", ' + $fontSize + ', [System.Drawing.FontStyle]::' + $fontStyle + ")"
                         }
                         else {
                              $value = $value -replace 'True', '$true' -replace 'False', '$false' -replace 'Color \[(\w+)\]', '$1'
                              if ($value -ne '$true' -and $value -ne '$false') {
                                   $value = '"' + $value + '"'
                              }
                         }
                         $Global:source += '$' + $container.Name + '.' + $propName + ' = ' + $value + $newline # Add Code for each property
                    }
               }
          }

          foreach ($event in $container.Tag) {
               # Grabs all events for the control and adds to the code
               if ($event -like "Add_*") {
                    $Global:source += '$' + $container.Name + '.' + $event + $newline
               }
          }

          if ($ctrlType -eq 'Form') {
               # Adds Form width and height properties to the code
               # --- Form ---
               $width = $container.ClientSize.Width
               $height = $container.ClientSize.Height
               $Global:source += '$' + $container.Name + '.ClientSize = New-Object System.Drawing.Size(' + $width + ', ' + $height + ')' + $newline
          }
          else {
               # Adds Control width, height, location, and parent container properties to the code
               # --- Other controls ---
               if ($width -ne 0 -and $height -ne 0) {
                    $Global:source += '$' + $container.Name + '.Size = New-Object System.Drawing.Size(' + $width + ', ' + $height + ')' + $newline
               }
               $Global:source += '$' + $container.Name + '.Location = New-Object System.Drawing.Point(' + $left + ', ' + $top + ')' + $newline
               $Global:source += '$' + $container.Parent.Name + '.Controls.Add($' + $container.Name + ')' + $newline
          }
          # --- Containers ---
          if ($ctrlType -eq 'Form' -or $ctrlType -eq 'GroupBox') {
               # Adds the containers to the code and then adds all controls attached to the Container.
               foreach ($ctrl in $container.Controls) {
                    Enumerate-SaveControls $ctrl
               }
          }
     }
     # Main process to build code for savefile
     # Add Global Properties to the code
     $newline = "`r`n"
     $Global:source = 'Add-Type -AssemblyName System.Windows.Forms' + $newline
     $Global:source += 'Add-Type -AssemblyName System.Drawing' + $newline
     $Global:source += '[Windows.Forms.Application]::EnableVisualStyles()' + $newline

     # Runs Enumerate-SaveControls Function to add form code.
     Enumerate-SaveControls $Global:frmDesign

     # Adds Show Dialog Code (Last Line in code) NOTE: Can be erased if dot-sourcing into a MainProgram file, keeps form code separate from main code (Future Feature)
     $Global:source += $newline + '[void]$' + $Global:frmDesign.Name + '.ShowDialog()' + $newline

     # Sets file name for the savefile
     [string]$filename = ''
     $filename = Open-FilenameDialog 'SaveFileDialog'
     if ($filename -notlike '') {
          $Global:source > $filename # Adds source code to the file and saves
     }
}

#--- Open Existing Form ---
function Open-DesignForm {
     # On Open Form Button Click, Searches for the form information in the file chosen and opens it into the Form Designer (Must be in the Standard Win form format)
     function Get-FormControlsWithProps($ctrl) {
          # Gets all Control info and their properties
          $pattern = '(.*)\$' + $ctrl.Name + '\.(?:(\w+)\s*=|(Add_[^\r\n]+))' # Matches patterns like $GroupBox0.word or $GroupBox0.Add_
          $allMatches = [regex]::matches($Global:source, $pattern)
          $tags = @()
          foreach ($item in $allMatches) {
               [string]$comment = $item.Groups[1]
               if (-not $comment.Contains('#')) {
                    $propName = [string]$item.Groups[2]
                    if ($propName) {
                         if ($propName -eq 'Location') {
                              $tags += @('Left', 'Top')
                         }
                         elseif ($propName -eq 'Size' -or $propName -eq 'ClientSize') {
                              $tags += @('Width', 'Height')
                         }
                         else {
                              $tags += $propName
                         }
                    }
                    $eventHandler = [string]$item.Groups[3]
                    if ($eventHandler) {
                         $tags += $eventHandler
                    }
               }
          }
          if ($tags -notcontains 'Name') {
               $tags += 'Name'
          }
          $ctrl.Tag = $tags
     }
     function Enumerate-LoadControls($container) {
          # Enumerates and loads all controls into the designer
          foreach ($ctrl in $container.Controls) {
               Get-FormControlsWithProps $ctrl
               $ctrlType = Get-CtrlType $ctrl
               if ($ctrlType -eq 'GroupBox') {
                    Enumerate-LoadControls $ctrl
               }
               if ($ctrlType -eq 'ComboBox') {
                    $ctrl.IntegralHeight = $false
               }
               elseif ($ctrlType -eq 'ListBox') {
                    $ctrl.IntegralHeight = $false
               }
               elseif ($ctrlType -ne 'WebBrowser') {
                    $ctrl.Cursor = 'SizeAll'
                    $ctrl.Add_PreviewKeyDown( {$_.IsInputKey = $true} )
                    $ctrl.Add_KeyDown( {Set-ResizeAndMoveWithKeyboard} )
                    $ctrl.Add_MouseDown( {Set-MouseDown} )
                    $ctrl.Add_MouseMove( {Set-MouseMove} )
                    $ctrl.Add_MouseUP( {Set-MouseUP} )
                    $ctrl.Add_Click( {Set-CurrentCtrl $this} )
               }
          }
     }
     # Main code to open a form
     $filename = Open-FilenameDialog 'OpenFileDialog'
     if ($filename -notlike '') {
          # Get the source of the file, and run a search to find the data for the form to be opened.
          $Global:source = get-content $filename | Out-String
          $pattern = '(.*)\$(\w+)\s*=\s*New\-Object\s+(System\.)?Windows\.Forms\.Form' # Pattern to find the Beggining of the Form Code by finding "New-Object System.Windows.Forms.Form"
          $allMatches = [regex]::matches($Global:source, $pattern)
          foreach ($item in $allMatches) {
               [string]$comment = $item.Groups[1]
               if (-not $comment.Contains('#')) {
                    $formName = $item.Groups[2]
               }
          }
          if ($formName) {
               # Use the form name to find the showdialog entry and remove the entry it will be loaded as code in the editor, we do not want it to run like a active form
               $find = '\$' + $formName + '\.Show(Dialog)?\(\)'
               $Global:source = $Global:source -replace $find, ''

               Invoke-Expression -Command $Global:source # Execute the Form code to open the form for editing

               try {
                    $Global:frmDesign = Get-Variable -ValueOnly $formName # Try to find the forms variable
               }
               catch {
               }
               if ($Global:frmDesign) {
                    # If the variable is positive, check to see if it is a form and load the properties and events to the Enumerate-LoadControls
                    Get-Variable | Where-Object {[string]$_.Value -like 'System.Windows.Forms.*'} | Where-Object {
                         try {
                              $_.Value.Name = $_.Name
                         }
                         catch {
                         }
                    }

                    Enumerate-LoadControls $Global:frmDesign
                    $Global:frmDesign.Name = $formName
                    Get-FormControlsWithProps $Global:frmDesign
                    $Global:frmDesign.Add_ResizeEnd( {Update-CurrentCrtlPropsGrid} )
                    $Global:frmDesign.Add_FormClosing( {Close-FrmDesign} )
                    $Global:frmDesign.Show()
                    $Global:currentCtrl = $Global:frmDesign
                    Update-ControlList
                    Update-AvailablePropsInCB
                    Update-CurrentCrtlPropsGrid
                    Update-AvailableEventsCB
                    Update-CurrentCtrlEventsGrid
                    Set-BtnsFrmDesignOpen
               }
               else {
                    # Error if forms variable not found in code
                    $message = "Can't find variable $" + $formName + "`nPlease open ONLY SOURCE OF FORM.`nExclude other code."
                    [System.Windows.Forms.MessageBox]::Show($message, 'Error opening existing Form', 'OK', 'Error')
               }
          }
          else {
               # Error if no form code is found
               $message = 'Your code does not contain a form!'
               [System.Windows.Forms.MessageBox]::Show($message, 'Error opening existing Form', 'OK', 'Error')
          }
     }
}

#endregion === Forms ===

#region === Main Application Window ===

# --- Build and Show Main Window ---

#frmPSFD
$frmPSFDMain = New-Object System.Windows.Forms.Form
$frmPSFDMain.ClientSize = New-Object System.Drawing.Size(549, 524)
$frmPSFDMain.FormBorderStyle = 'Fixed3D'
$frmPSFDMain.MaximizeBox = $false
$frmPSFDMain.Text = 'PowerShell WinForm Designer ' + $Version

#bNewForm
$bNewForm = New-Object System.Windows.Forms.Button
$bNewForm.Location = New-Object System.Drawing.Point(12, 12)
$bNewForm.Size = New-Object System.Drawing.Size(88, 23)
$bNewForm.Text = "New Form"
$bNewForm.Add_Click( {Create-NewFrmDesign} )
$frmPSFDMain.Controls.Add($bNewForm)

#bOpenForm
$bOpenForm = New-Object System.Windows.Forms.Button
$bOpenForm.Location = New-Object System.Drawing.Point(114, 12)
$bOpenForm.Size = New-Object System.Drawing.Size(88, 23)
$bOpenForm.Text = "Open Form"
$bOpenForm.Add_Click( {Open-DesignForm} )
$frmPSFDMain.Controls.Add($bOpenForm)

#bSaveForm
$bSaveForm = New-Object System.Windows.Forms.Button
$bSaveForm.Location = New-Object System.Drawing.Point(450, 12)
$bSaveForm.Size = New-Object System.Drawing.Size(88, 23)
$bSaveForm.Text = "Save Form"
$bSaveForm.Enabled = $false
$bSaveForm.Add_Click( {Save-FrmDesign} )
$frmPSFDMain.Controls.Add($bSaveForm)

# --- Controls Section ---

#gbControls
$gbControls = New-Object Windows.Forms.GroupBox
$gbControls.Location = New-Object System.Drawing.Point(12, 41)
$gbControls.Size = New-Object System.Drawing.Size(260, 472)
$gbControls.Text = "Controls:"
$frmPSFDMain.Controls.Add($gbControls)

#cbAddControl
$cbAddControl = New-Object Windows.Forms.ComboBox
$cbAddControl.Location = New-Object System.Drawing.Point(6, 14)
$cbAddControl.Size = New-Object System.Drawing.Size(182, 21)
$cbAddControl.Items.AddRange(@("Button", "CheckBox", "ComboBox", "DataGridView", "DateTimePicker", "GroupBox", "Label", "ListBox", "ListView", "RadioButton", "PictureBox", "RichTextBox", "TextBox", "TreeView"))
$gbControls.Controls.Add($cbAddControl)

#bAddControl
$bAddControl = New-Object System.Windows.Forms.Button
$bAddControl.Location = New-Object System.Drawing.Point(196, 14)
$bAddControl.Size = New-Object System.Drawing.Size(58, 23)
$bAddControl.Text = "Add"
$bAddControl.Add_Click( {Add-Control} )
$bAddControl.Enabled = $false
$gbControls.Controls.Add($bAddControl)

#bRemoveControl
$bRemoveControl = New-Object System.Windows.Forms.Button
$bRemoveControl.Location = New-Object System.Drawing.Point(196, 40)
$bRemoveControl.Size = New-Object System.Drawing.Size(58, 23)
$bRemoveControl.Text = "Remove"
$bRemoveControl.Enabled = $false
$bRemoveControl.Add_Click( {Remove-CurrentCtrl} )
$gbControls.Controls.Add($bRemoveControl)

#labTooltipC
$labTooltipC = New-Object System.Windows.Forms.Label
$labTooltipC.Location = New-Object System.Drawing.Point(6, 44)
$labTooltipC.Size = New-Object System.Drawing.Size(180, 16)
$labTooltipC.Text = "Add or Remove Controls here."
$labTooltipC.Enabled = $false
$gbControls.Controls.Add($labTooltipC)

#dgvControls
$dgvControls = New-Object System.Windows.Forms.DataGridView
$dgvControls.Location = New-Object System.Drawing.Point(6, 70)
$dgvControls.Size = New-Object System.Drawing.Size(248, 396)
$null = $dgvControls.Columns.Add("", "Name")
$null = $dgvControls.Columns.Add("", "Type")
$null = $dgvControls.Columns.Add("", "ParentControl")
$null = $dgvControls.Columns.Add("", "LinkToControl")
$dgvControls.AutoSizeColumnsMode = 'Fill'
$dgvControls.Columns[0].ReadOnly = $true
$dgvControls.Columns[1].ReadOnly = $true
$dgvControls.Columns[2].ReadOnly = $true
$dgvControls.Columns[3].Visible = $false
$dgvControls.ColumnHeadersHeightSizeMode = 'DisableResizing'
$dgvControls.RowHeadersVisible = $false
$dgvControls.MultiSelect = $false
$dgvControls.ScrollBars = 'Vertical'
$dgvControls.SelectionMode = 'FullRowSelect'
$dgvControls.AllowUserToResizeRows = $false
$dgvControls.AllowUserToAddRows = $false
$dgvControls.Add_CellClick( {Set-CurrentCtrl $dgvControls.CurrentRow.Index} )
$gbControls.Controls.Add($dgvControls)

#--- Properties Section ---

#gbProps
$gbProps = New-Object Windows.Forms.GroupBox
$gbProps.Location = New-Object System.Drawing.Point(278, 41)
$gbProps.Size = New-Object System.Drawing.Size(260, 350)
$gbProps.Text = 'Properties:'
$frmPSFDMain.Controls.Add($gbProps)

#cbAddProp
$cbAddProp = New-Object Windows.Forms.ComboBox
$cbAddProp.Location = New-Object System.Drawing.Point(6, 14)
$cbAddProp.Size = New-Object System.Drawing.Size(182, 21)
$gbProps.Controls.Add($cbAddProp)

#bAddProp
$bAddProp = New-Object System.Windows.Forms.Button
$bAddProp.Location = New-Object System.Drawing.Point(196, 14)
$bAddProp.Size = New-Object System.Drawing.Size(58, 23)
$bAddProp.Text = "Add"
$bAddProp.Add_Click( {Add-Property $cbAddProp.Text})
$bAddProp.Enabled = $false
$gbProps.Controls.Add($bAddProp)

#bDelProp
$bRemoveProp = New-Object System.Windows.Forms.Button
$bRemoveProp.Location = New-Object System.Drawing.Point(196, 40)
$bRemoveProp.Size = New-Object System.Drawing.Size(58, 23)
$bRemoveProp.Text = "Remove"
$bRemoveProp.Add_Click( {Remove-Property $dgvProps.CurrentRow.Cells[0].Value} )
$bRemoveProp.Enabled = $false
$gbProps.Controls.Add($bRemoveProp)

#labTooltip
$labTooltipP = New-Object System.Windows.Forms.Label
$labTooltipP.Location = New-Object System.Drawing.Point(6, 44)
$labTooltipP.Size = New-Object System.Drawing.Size(180, 16)
$labTooltipP.Text = "Add or Remove Properties here."
$labTooltipP.Enabled = $false
$gbProps.Controls.Add($labTooltipP)

#dgvProps
$dgvProps = New-Object System.Windows.Forms.DataGridView
$dgvProps.Location = New-Object System.Drawing.Point(6, 70)
$dgvProps.Size = New-Object System.Drawing.Size(248, 274)
$null = $dgvProps.Columns.Add("", "Property")
$null = $dgvProps.Columns.Add("", "Value")
$dgvProps.Columns[0].ReadOnly = $true
$dgvProps.AutoSizeColumnsMode = 'Fill'
$dgvProps.ColumnHeadersHeightSizeMode = 'DisableResizing'
$dgvProps.RowHeadersVisible = $false
$dgvProps.AllowUserToResizeRows = $false
$dgvProps.AllowUserToAddRows = $false
$dgvProps.Add_CellEndEdit( {Set-PropGridCellEndEdit} )
$gbProps.Controls.Add($dgvProps)

#--- Events Section ---

#gbEvents
$gbEvents = New-Object System.Windows.Forms.GroupBox
$gbEvents.Text = "Event Handlers:"
$gbEvents.Size = New-Object System.Drawing.Size(260, 114)
$gbEvents.Location = New-Object System.Drawing.Point(278, 398)
$frmPSFDMain.Controls.Add($gbEvents)

#cbEvents
$cbAddEvent = New-Object System.Windows.Forms.ComboBox
$cbAddEvent.Size = New-Object System.Drawing.Size(182, 21)
$cbAddEvent.Location = New-Object System.Drawing.Point(6, 14)
$gbEvents.Controls.Add($cbAddEvent)

#bAddEvent
$bAddEvent = New-Object System.Windows.Forms.Button
$bAddEvent.Text = "Add"
$bAddEvent.Size = New-Object System.Drawing.Size(58, 23)
$bAddEvent.Location = New-Object System.Drawing.Point(197, 14)
$bAddEvent.Add_Click( {Add-Event} )
$gbEvents.Controls.Add($bAddEvent)

#dgvEvents
$dgvEvents = New-Object System.Windows.Forms.DataGridView
$dgvEvents.Size = New-Object System.Drawing.Size(248, 66)
$dgvEvents.Location = New-Object System.Drawing.Point(6, 43)
$null = $dgvEvents.Columns.Add("", "Event")
$dgvEvents.Columns[0].Width = 245
$dgvEvents.ColumnHeadersVisible = $false
$dgvEvents.RowHeadersVisible = $false
$dgvEvents.AllowUserToResizeRows = $false
$dgvEvents.AllowUserToAddRows = $false
$dgvEvents.ScrollBars = 'Vertical'
$gbEvents.Controls.Add($dgvEvents)

#--- Start Form ---

[void]$frmPSFDMain.ShowDialog()

#endregion === Main Application Window ===
