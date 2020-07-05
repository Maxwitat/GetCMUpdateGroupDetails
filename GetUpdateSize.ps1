##************************************************************************************************************************
##    UpdateGroupDetails.ps1
##   
##    Frank Maxwitat, July 05, 2020, version 1.0
##    
##    Return details the content and size information for an update group as html report
##
##    Usage: UpdateGroupDetails -CMDBName 'CM_[SiteCode]' -CMDBServerName '[DatabaseServer]' -UpdateGroupName '[Group Name]'
##    Example: UpdateGroupDetails -CMDBName 'CM_P01' -CMDBServerName 'SVRP01' -UpdateGroupName '06_2020 Update'
##
##************************************************************************************************************************
    
function UpdateGroupDetails
{
    param (
    [Parameter(Mandatory=$true)]
    $CMDBName,
    [Parameter(Mandatory=$true)]
    $CMDBServerName,
    [Parameter(Mandatory=$true)]
    $UpdateGroupName,
    [Parameter(Mandatory=$false)]
    $OutputFolder
    )

    $Report = If ($OutputFolder) {$OutputFolder + '\' + $UpdateGroupName + '.html'} Else {$PSScriptRoot + '\' + $UpdateGroupName + '.html'}
    if(Test-Path($Report)) {Remove-Item -Path $Report -Force}
    $ReportTitle = "Content details for Update Group " + $UpdateGroupName

    $header = @"
   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>$Title</title>
    <style type="text/css">
    <!--
    body {
            font: 100%/1.4 Verdana, Arial, Helvetica, sans-serIf;
            background: #FFFFFF;
            margin: 0;
            padding: 0;
            color: #000;
         }
    .container {
            width: 100%;
            margin: 0 auto;
            }
    h1 {
            font-size: 18px;
        }
    h2 {
            color: #FFF;
            padding: 0px;
            margin: 0px;
            font-size: 14px;
            background-color: #006400;
        }
    h3 {
            color: #FFF;
            padding: 0px;
            margin: 0px;
            font-size: 14px;
            background-color: #191970;
        }
    h4 {
            color: #348017;
            padding: 0px;
            margin: 0px;
            font-size: 10px;
            font-style: italic;
        }
    .header {
            text-align: center;
        }
    .container table {
            width: 100%;
            font-family: Verdana, Geneva, sans-serIf;
            font-size: 12px;
            font-style: normal;
            font-weight: bold;
            font-variant: normal;
            text-align: center;
            border: 0px solid black;
            padding: 0px;
            margin: 0px;
        }
    td {
            font-weight: normal;
            border: 1px solid grey;
            width='25%'
        }
    th {
            font-weight: bold;
            border: 1px solid grey;
            text-align: center;
        }
    -->
    </style></head>
    <body>
    <div class="container">
    <div class="content">	
"@
    Add-Content "$Report" $header	
	$RptHeaderSME1 = @"
	<table width='100%'><tbody>
	<tr bgcolor = '$HeaderBGColor'> <td align='center'> <b> 
	<Font color = 'white'> $ReportTitle </Font>
	</b> </td> </tr>
	</table>
"@
    Add-Content $Report $RptHeaderSME1    

    $objConnection = New-Object -comobject ADODB.Connection		
	$constr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Initial Catalog=$CMDBName;Data Source=$CMDBServerName"
	$objConnection.Open($constr)
	$objConnection.CommandTimeout = 0
	# *********** Check If connection is open *******************
	If($objConnection.state -eq 0)
	{
		Write-Host "Error: ConfigMgr DB ServerName or ConfigMgr DB Name are not correct your your account does not have sufficient permissions" -ForegroundColor Yellow -BackgroundColor DarkRed
		Exit 1
	}

$strSQL=@"
    SELECT distinct v_UpdateInfo.Title, v_UpdateInfo.CI_UniqueID
    FROM v_UpdateInfo 
	INNER JOIN v_CIAssignmentToCI ON v_UpdateInfo.CI_ID = v_CIAssignmentToCI.CI_ID
	INNER JOIN v_CIRelation ON v_UpdateInfo.CI_ID = v_CIRelation.ToCIID 
	INNER JOIN v_AuthListInfo ON v_CIRelation.FromCIID = v_AuthListInfo.CI_ID 
	where v_AuthListInfo.Title like '$UpdateGroupName'
    ORDER BY v_UpdateInfo.Title
"@
            
$rptheader=@"
    <table width='100%'><tbody>
	<tr bgcolor=$TableHeaderBGColor> <td> <b> <Font color = 'white'> Update Details $UpdateGroupName </Font> </b> </td> </tr>
    </table>
    <table width='100%' border = 0 > <tbody>
	<tr bgcolor=$TableHeaderRowBGColor>
    <td width='30%'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:black'>Title</span></b></td>
    <td width='7%'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:black'>Severity</span></b></td>
    <td width='7%'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:black'>Article ID</span></b></td>
    <td width='23%'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:black'>Folder(s)</span></b></td>
    <td width='23%'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:black'>File(s)</span></b></td>
    <td width='6%'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:black'>Size (MB)</span></b></td>
    <td width='4%'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:black'>Parts</span></b></td>     
	</tr>
    </table>
"@
    Add-Content "$Report" $rptheader 

    $TotalSizeInMB = 0
    $objRecordset = New-Object -comobject ADODB.Recordset
    $objRecordset.Open($strSQL,$objConnection)         
	$objRecordset.MoveFirst()
	$rows=$objRecordset.RecordCount
	do 
	{
		$UpdateTitle = $objRecordset.Fields.Item(0).Value
		$UpdateCI_UniqueID = $objRecordset.Fields.Item(1).Value			             

        Write-Host "Adding Update " $UpdateTitle " Unique ID is " $UpdateCI_UniqueID  -BackgroundColor DarkGreen
        [string]$UID = $UpdateTitle

		$objRecordset2 = New-Object -comobject ADODB.Recordset		        
$strSQL = @" 
        select UI.Title, CASE(Severity) When 2 Then 'Low' When 6 Then 'Moderate' When 8 Then 'Important' When 10 Then 'Critical' Else 'NA' End as 'Severity', 
        UI.DateLastModified, ArticleID, PKG.PkgSourcePath, CNT.Content_UniqueID from v_UpdateInfo as UI
        inner join v_Update_ComplianceSummary as UCS on UCS.CI_ID = UI.CI_ID --use this if you want compliance details per update
        inner join vUpdateToPkg as VUTP on VUTP.CI_ID = UI.CI_ID
        inner join v_CIToContent AS CNT ON CNT.CI_ID = VUTP.CI_ID
        inner join v_Package AS PKG ON PKG.PackageID=VUTP.PkgID
        where UI.CI_UniqueID like '$UpdateCI_UniqueID'
"@                
        $objRecordset2.Open($strSQL,$objConnection)
        $objRecordset2.MoveFirst()
		$UpdateSizeInMB = 0
        $Parts = 0
        $Folder = ''; $Files = '';
		if($objRecordset2.Fields.Item(0).Value)
        {
            do #there can be more than one folder, therefore I sum up the sizes (alternatively, you may add one line per folder)
            {		                
                $Title = $objRecordset2.Fields.Item(0).Value
		        $UpdateSeverity = $objRecordset2.Fields.Item(1).Value
		        $DateLastModified  = $objRecordset2.Fields.Item(2).Value
		        $ArticleID = $objRecordset2.Fields.Item(3).Value
                $PkgPath = $objRecordset2.Fields.Item(4).Value
                $PkgUID = $objRecordset2.Fields.Item(5).Value

                $UpdatePath = $PkgPath + '\' + $PkgUID
                $Folder += '<p>'+ $PkgUID + '</p>'
                    
                #$UpdatePath
                $Parts++ 
                $FolderSize = "{0:N2}" -f ((Get-ChildItem -path $UpdatePath -recurse | Measure-Object -property length -sum ).sum /(1024*1024))
                $Files += '<p>'+ (Get-ChildItem -path $UpdatePath).Name + '</p>'
                $UpdateSizeInMB += $FolderSize
                $TotalSizeInMB += $FolderSize
$Rpt=@"
                <table width='100%' border = 0 > <tbody>
	            <tr align='Left'>
                <td width='30%' align='left'>$Title</td>
                <td width='7%' align='left'>$UpdateSeverity</td>
                <td width='7%' align='left'>$ArticleID</td>
                <td width='23%' align='left'>$Folder</td>
                <td width='23%' align='left'>$Files</td>
                <td width='6%' align='left'>$UpdateSizeInMB</td>
                <td width='4%' align='left'>$Parts</td>
	            </tr>
                </table>
"@
                        
                $objRecordset2.MoveNext()
            } 
		    until ($objRecordset2.EOF -eq $TRUE)
            Add-Content "$Report" $Rpt
        }
        if($objRecordset2)
        {
            $objRecordset2.Close()
        }
        $objRecordset.MoveNext()
	} 
	until ($objRecordset.EOF -eq $TRUE)		 
    if($objConnection)
    {
        $objRecordset.Close()
        $objConnection.Close()
    }
    
    $RptFooter1 = @"
    <table width='100%' bgcolor = '$FooterBGColor'><tbody>
   	<tr> <td align='center'> <b> <Font color = 'white'> Total size: $TotalSizeInMB MB - Content path: $PkgPath </Font> </b> </td> </tr>
    </table>
"@
    Add-Content $Report $RptFooter1
}

UpdateGroupDetails -CMDBName 'CM_P01' -CMDBServerName 'SVRP01' -UpdateGroupName '06_2020 Update'