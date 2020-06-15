# -------------------------------
# BI Deployment Scripts Generator
# -------------------------------

add-type -AssemblyName System.Data.OracleClient

$oraUser    = 'DWH_OPR_NL'
$oraPass    = 'dwh_opr_nl_d'
$dataSource = 'etld01pw'

$connection_string = "User Id=$oraUser;Password=$oraPass;Data Source=$dataSource"
$stm_release = "SELECT r.release_nr
                     , r.revision_label
                     , r.environment
                     , r.root_directory
                     , r.checkout_directory
                     , r.deploy_file
                     , r.log_file
                  FROM isg_v_release r 
                 WHERE r.status = 'Ready for release'
                 ORDER BY r.revision_label
                        , r.environment"

$stm_script = "SELECT statement_sql
                 FROM isg_statement
                WHERE release_nr = :an_release_nr
                  AND statement_file = 'install'
                ORDER BY statement_nr"

try{
    $con = New-Object System.Data.OracleClient.OracleConnection($connection_string)
    $con.Open()

#########################

    $cmd_release = $con.CreateCommand()
    $cmd_release.CommandText = $stm_release

#########################

    $cmd_script = $con.CreateCommand()
    $cmd_script.CommandText = $stm_script

#########################

<#
    $cmd_get_label = $con.CreateCommand()
    $cmd_get_label.CommandType = [System.Data.CommandType]::StoredProcedure
    $cmd_get_label.CommandText = "pck_dwh_isg_rsch8.get_label"

    $cmd_get_label.Parameters.Add("an_release_nr", [System.Data.OracleClient.OracleType]::Number) | out-null;
    $cmd_get_label.Parameters["an_release_nr"].Direction = [System.Data.ParameterDirection]::InputOutput;

    $cmd_get_label.Parameters.Add("as_label", [System.Data.OracleClient.OracleType]::VarChar) | out-null;
    $cmd_get_label.Parameters["as_label"].Direction = [System.Data.ParameterDirection]::InputOutput;
#>
#########################
    
    $cmd_ins_item = $con.CreateCommand()
    $cmd_ins_item.CommandType = [System.Data.CommandType]::StoredProcedure
    $cmd_ins_item.CommandText = "pck_dwh_isg_rsch8.insert_checked_out_object"

    $cmd_ins_item.Parameters.Add("an_release_nr", [System.Data.OracleClient.OracleType]::Number) | out-null;
    $cmd_ins_item.Parameters["an_release_nr"].Direction = [System.Data.ParameterDirection]::Input;

    $cmd_ins_item.Parameters.Add("an_item_nr", [System.Data.OracleClient.OracleType]::Number) | out-null;
    $cmd_ins_item.Parameters["an_item_nr"].Direction = [System.Data.ParameterDirection]::Input;
    
    $cmd_ins_item.Parameters.Add("as_file_name", [System.Data.OracleClient.OracleType]::VarChar) | out-null;
    $cmd_ins_item.Parameters["as_file_name"].Direction = [System.Data.ParameterDirection]::Input;

######################### 

    $cmd_isg_update = $con.CreateCommand()
    $cmd_isg_update.CommandType = [System.Data.CommandType]::StoredProcedure
    $cmd_isg_update.CommandText = "pck_dwh_isg_rsch8.isg_update"

    $cmd_isg_update.Parameters.Add("an_release_nr", [System.Data.OracleClient.OracleType]::Number) | out-null;
    $cmd_isg_update.Parameters["an_release_nr"].Direction = [System.Data.ParameterDirection]::Input;

######################### 

    $cmd_end_release = $con.CreateCommand()
    $cmd_end_release.CommandType = [System.Data.CommandType]::StoredProcedure
    $cmd_end_release.CommandText = "pck_dwh_isg_rsch8.end_release"

    $cmd_end_release.Parameters.Add("an_release_nr", [System.Data.OracleClient.OracleType]::Number) | out-null;
    $cmd_end_release.Parameters["an_release_nr"].Direction = [System.Data.ParameterDirection]::Input;

    $cmd_end_release.Parameters.Add("as_log_data", [System.Data.OracleClient.OracleType]::Clob) | out-null;
    $cmd_end_release.Parameters["as_log_data"].Direction = [System.Data.ParameterDirection]::Input;

######################### 
<#
    $cmd_install_script = $con.CreateCommand()
    $cmd_install_script.CommandType = [System.Data.CommandType]::StoredProcedure
    $cmd_install_script.CommandText = "pck_dwh_isg_rsch8.install_script"

    $cmd_install_script.Parameters.Add("as_install_script", [System.Data.OracleClient.OracleType]::Clob) | out-null;
    $cmd_install_script.Parameters["an_release_nr"].Direction = [System.Data.ParameterDirection]::ReturnValue;

    $cmd_install_script.Parameters.Add("an_release_nr", [System.Data.OracleClient.OracleType]::Number) | out-null;
    $cmd_install_script.Parameters["an_release_nr"].Direction = [System.Data.ParameterDirection]::InputOutput;
#>
#########################

    Write-Host "Fetching release data"
    $rdr_release = $cmd_release.ExecuteReader()

    while ($rdr_release.Read()) {

       $release_nr         = $rdr_release.GetDecimal(0)
       $revision_label     = $rdr_release.GetString(1)
       $environment        = $rdr_release.GetString(2)
       $root_directory     = $rdr_release.GetString(3)
       $checkout_directory = $rdr_release.GetString(4)
       $deploy_file        = $rdr_release.GetString(5)
       $log_file           = $rdr_release.GetString(6)

<#

       Write-Host "About to checkout items from StarTeam server ..."

       # remove previously checked out files
       if (Test-Path $checkout_directory) { Remove-Item -Force -LiteralPath "$checkout_directory" -Recurse }

       # get files from StarTeam by revision label, will create directory "Datawarehouse"
       Write-Host "connecting to StarTeam server ..."
       stcmd connect prut1@nlrtdsrv4star01:49998  | out-null;
       Write-Host "setting project and folder ..."
       stcmd set project="Oracle" folderHierarchy="Datawarehouse"  | out-null;
       Write-Host "checking out files with label `"$revision_label`" ..."
       stcmd co -vl "$revision_label" -rp "$checkout_directory" -frp -is -o  | out-null;
       Write-Host "disconnecting ..."
       stcmd disconnect  | out-null;

       # exit if nothing was checked out
       if (!(Test-Path $checkout_directory)) {
          Write-Host "There were no files to check out with label `"$revision_label`".
          [Abort]"
          exit
       }
#>
       $an_item_nr = 0
       # loop through source files
       Get-ChildItem -Path $checkout_directory -Recurse -File | ForEach-Object {
         $an_item_nr++
         $cmd_ins_item.Parameters["an_release_nr"].Value = $release_nr
         $cmd_ins_item.Parameters["an_item_nr"].Value = $an_item_nr
         $cmd_ins_item.Parameters["as_file_name"].Value = $_.FullName
<#
         Write-Host $cmd_ins_item.Parameters["an_release_nr"].Value
         Write-Host $cmd_ins_item.Parameters["an_item_nr"].Value
         Write-Host $cmd_ins_item.Parameters["as_file_name"].Value
#>
         Write-Host "Inserting to be released item names in database"
         $cmd_ins_item.ExecuteNonQuery() | out-null;

         # replace backslashes by slash and remove root part of name
         # $name = $_.FullName.Replace("\", "/").Replace("$root_directory/", "")
         
       }

       Write-Host "Prepare statements in database "

       $cmd_isg_update.Parameters["an_release_nr"].Value = $release_nr
       $cmd_isg_update.ExecuteNonQuery() | out-null;

       Write-Host "Getting statements to add to release script"

       $cmd_script.Parameters.Add("an_release_nr", $release_nr)
       $rdr_script = $cmd_script.ExecuteReader()

       Set-Content $deploy_file ''
       
       Write-Host "Create deploy file using statements read from database."
       while ($rdr_script.Read()) {
          $statement_sql = $rdr_script.GetString(0)
          # Write-Host $statement_sql
          Add-Content $deploy_file $statement_sql
       }

       Write-Host "Run sqlplus to deploy items in database."
       sqlplus -s /nolog @$deploy_file $statement_sql

       Write-Host "Reading sqlplus logfile in variable."
       $log_data = Get-Content $log_file -Raw

       Write-Host "Adding sqlplus logfile to release in database."

       $cmd_end_release.Parameters["an_release_nr"].Value = $release_nr
       $cmd_end_release.Parameters["as_log_data"].Value = $log_data
       $cmd_end_release.ExecuteNonQuery() | out-null;
      
   }
} catch {
    Write-Error ("Database Exception: {0}`n{1}” -f `
        $con.ConnectionString, $_.Exception.ToString())
} finally{
    if ($con.State -eq ‘Open’) { $con.close() }
}

