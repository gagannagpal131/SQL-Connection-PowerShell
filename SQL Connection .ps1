#Variable Stores names of the installed programs in the local system
$insprog = gwmi win32_installedwin32program | select name

#server details of MSSQL 
#Setting up connection between MSSQL and PowerShell
$SQLServer = "WINDOWS-1\SQLDEV"
$SQLDBName = "test"
$uid ="WINDOWS-1\*******"
#$pwd = "d&rh=oI%MkPSio%"
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True; User ID = $uid; Password = $pwd;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand

#Pushing data into the server
foreach($i in $insprog){
$SqlQuery = "INSERT INTO Apps (Apps) VALUES ('$i');"
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
}

#fetching data from the server and pushing data into a .csv file
$SqlQuery = "SELECT * FROM Apps;"
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$DataSet.Tables[0] | out-file "C:\Users\******\*****\*****\Apps.csv"