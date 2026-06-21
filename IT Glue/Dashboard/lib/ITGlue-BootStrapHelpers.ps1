<#
    ITGlue-BootStrapHelpers.ps1

    Bootstrap-3 panel helper functions for rendering HTML inside IT Glue Flexible Assets.

    SOURCE / ATTRIBUTION
    --------------------
    These two functions (New-BootstrapSinglePanel, New-BootstrapInfoPanel) are vendored
    verbatim from Gavin Stone's "ITGlue-Helper" project:
        https://github.com/gavsto/ITGlue-Helper
        IT Glue BootStrap Panel Generator/ITGlue-BootStrapHelpers.ps1
    Original author: Gavin Stone (2020-08-17). Licensed GPL-3.0.

    The upstream demo/example block at the bottom of the original file has been removed;
    only the reusable functions are kept. Behaviour is unchanged.

    WHY THESE WORK IN IT GLUE
    -------------------------
    IT Glue's own web UI is built on Bootstrap 3, so the `panel panel-success` /
    `panel-danger` / `label-*` / `alert-*` classes emitted here inherit IT Glue's own
    stylesheet when the HTML is stored in a Flexible Asset "Textbox" trait. No external CSS
    is required. IT Glue sanitises stored HTML (strips <script>/<style>/iframes/handlers)
    but preserves structural tags, class attributes, inline styles, and <a href> links.
#>

function New-BootstrapSinglePanel {
<#
.SYNOPSIS
Create a single panel for use in an IT Glue Flexible Asset
.DESCRIPTION
Takes a number of parameters and returns HTML that can be inserted in to IT Glue. Panel sizes correspond to BootStrap sizes, which are 12 wide.
.PARAMETER PanelShading
Valid options are 'active', 'success', 'info', 'warning', 'danger', 'blank' - these correspond to BootStrap 3 colours
.PARAMETER PanelTitle
Takes a title, can either be text or image. This is the title at the of the Panel
.PARAMETER PanelContent
The main content of the panel
.PARAMETER ContentAsBadge
Presents the content as a BootStrap Badge
.PARAMETER PanelAdditionalDetail
Used to add additional items underneath the main content
.OUTPUTS
BootStrap 3 compliant HTML which can be used inside an IT Glue Flexible Asset
.NOTES
Version:        1.0.0
Author:         Gavin Stone
Creation Date:  2020-08-17
Purpose/Change: Initial script development

.EXAMPLE
New-BootstrapSinglePanel -PanelShading "success" -PanelTitle "<img src='https://www.xyz.com/image.png'>" -ContentAsBadge -PanelSize 3 -PanelContent "<b>Active</b>" -PanelAdditionalDetail "$($YourVariable) Licenses"
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('active', 'success', 'info', 'warning', 'danger', 'blank')]
        [string]$PanelShading,

        [Parameter(Mandatory)]
        [string]$PanelTitle,

        [Parameter(Mandatory)]
        [string]$PanelContent,

        [switch]$ContentAsBadge,

        [string]$PanelAdditionalDetail,

        [Parameter(Mandatory)]
        [int]$PanelSize = 3
    )

    if ($PanelShading -ne 'Blank') {
        $PanelStart = "<div class=`"col-sm-$PanelSize`"><div class=`"panel panel-$PanelShading`">"
    }
    else {
        $PanelStart = "<div class=`"col-sm-$PanelSize`"><div class=`"panel`">"
    }

    $PanelTitle = "<div class=`"panel-heading`"><h3 class=`"panel-title text-center`">$PanelTitle</h3></div>"


    if ($PSBoundParameters.ContainsKey('ContentAsBadge')) {
        $PanelContent = "<div class=`"panel-body text-center`"><h4><span class=`"label label-$PanelShading`">$PanelContent</span></h4>$PanelAdditionalDetail</div>"
    }
    else {
        $PanelContent = "<div class=`"panel-body text-center`"><h4>$PanelContent</h4>$PanelAdditionalDetail</div>"
    }
    $PanelEnd = "</div></div>"
    $FinalPanelHTML = "{0}{1}{2}{3}" -f $PanelStart, $PanelTitle, $PanelContent, $PanelEnd
    return $FinalPanelHTML
}

function New-BootstrapInfoPanel {
<#
.SYNOPSIS
Create an info panel for use in an IT Glue Flexible Asset, the info panel consists of BootStrap Alerts
.DESCRIPTION
Takes a number of parameters and returns HTML that can be inserted in to IT Glue. Panel sizes correspond to BootStrap sizes, which are 12 wide.
This is designed to take an array of objects so you can add multiple
.PARAMETER PanelShading
Valid options are 'active', 'success', 'info', 'warning', 'danger', 'blank' - these correspond to BootStrap 3 colours
.PARAMETER PanelTitle
Takes a title, can either be text or image. This is the title at the of the Panel
.PARAMETER PanelContent
The main content of the panel. This should be an object that is built up as so:
$CustomInfoPanel = [PSCustomObject]@()
    $CustomInfoPanel += @{
        Shading = "success"
        AlertText = "There is no Server 2008 or SBS machines at this client"
    }
    $CustomInfoPanel += @{
        Shading = "danger"
        AlertText = "There are $Server2008Count Server 2008 or SBS insecure machines active"
    }
As many of these as needed can be added

.OUTPUTS
BootStrap 3 compliant HTML which can be used inside an IT Glue Flexible Asset
.NOTES
Version:        1.0.0
Author:         Gavin Stone
Creation Date:  2020-08-17
Purpose/Change: Initial script development

.EXAMPLE
New-BootstrapInfoPanel -PanelShading "info" -PanelContent $CustomInfoPanel -PanelSize 6 -PanelTitle "<img src='https://www.xyz.com/image.png'>"
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('active', 'success', 'info', 'warning', 'danger', 'blank')]
        [string]$PanelShading,

        [Parameter(Mandatory)]
        [string]$PanelTitle,

        [Parameter(Mandatory)]
        [pscustomobject[]]$PanelContent,

        [Parameter(Mandatory)]
        [int]$PanelSize = 3
    )

    if ($PanelShading -ne 'Blank') {
        $PanelStart = "<div class=`"col-sm-$PanelSize`"><div class=`"panel panel-$PanelShading`">"
    }
    else {
        $PanelStart = "<div class=`"col-sm-$PanelSize`"><div class=`"panel`">"
    }

    if (-not([string]::IsNullOrEmpty($PanelTitle))) {
        $PanelTitle = "<div class=`"panel-heading`"><h3 class=`"panel-title text-center`">$PanelTitle</h3></div>"
    }
    else {
        $PanelTitle = ""
    }

    $FinalPanelContent = "<div class=`"panel-body text-center`">"
    foreach ($item in $PanelContent) {
        $FinalPanelContent = "$FinalPanelContent<div class=`"alert alert-$($item.shading)`" role=`"alert`">$($item.AlertText)</div>"
    }

    $FinalPanelEnd = "</div>"


    $PanelEnd = "</div></div>"
    $FinalPanelHTML = "{0}{1}{2}{3}{4}" -f $PanelStart, $PanelTitle, $FinalPanelContent, $FinalPanelEnd, $PanelEnd
    return $FinalPanelHTML
}
