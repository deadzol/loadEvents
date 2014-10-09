#add to distributed com users & event log readers local groups
#wmimgmt -> wmi control -> security tab -> cimv2 -> security -> add user & grant: enable account & remote enable -> advanced -> edit -> this namespace and subnamespces
#dcomcnfg -*-> wmi -> security -> Access Permission -> Customize -> Edit -> add user & grant remote access

$server = "<mssql server>"
$database = "<db>"
$table = "<table>"
$servers = @("a", "b", "c")

$connection = New-Object System.Data.SQLClient.SQLconnection
$connection.connectionString = "server='$Server';database='$database';trusted_connection=true;"

foreach($server in $servers) {
	$connection.open()
	$command = New-Object System.Data.SQLClient.SQLcommand
	$command.connection = $connection
	$sql =  "select max(recordnumber) as count from eventlogs where servername like '$($server)%'"
	write-host "sql:" $sql
	$command.commandtext = $sql
	$sqlReader = $command.executereader()
	$ret = $sqlReader.Read()
	$lastEvent = $sqlReader["count"]
	#$sqlReader.close
	$connection.close()
	write-host "last:" $lastEvent
	if ($lastEvent -eq [System.DBNull]::Value) { $lastEvent = 0 }
	$query = "SELECT * FROM Win32_NTLogEvent WHERE EventCode != 5156 and Logfile = 'Security' and RecordNumber > $($lastEvent)"  #the eventcode exclusion is bad, need to remove this after built out.
	write-host "query: " $query
	$events = Get-WmiObject -Query $query -computername $server
	foreach($i in $events) {
		$connection.open()
		$command = New-Object System.Data.SQLClient.SQLcommand
		$command.connection = $connection
		$sql = "insert into $table ( ServerName, LogFile, EventID, CategoryString, EventType, Message, RecordNumber, 	SourceName, TimeGeneratedString, TimeGenerated, Type, UserName ) 
					values ( '$($i.ComputerName)', '$($i.Logfile)', '$($i.EventCode)', '$($i.CategoryString)', '$($i.TypeEvent)', '$($i.Message)', '$($i.RecordNumber)', '$($i.SourceName)', '$($i.TimeGenerated)', '', '$($i.Type)', '$($i.UserName)');"
		#write-host $sql
		$command.commandtext = $sql
		$ret = $command.executenonquery()
		$connection.close()
	}
	$connection.close()
}
