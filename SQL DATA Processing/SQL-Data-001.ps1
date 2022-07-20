$dataSource='dteksan1.dtek.com'
$database='DTEK_scratch_test'

$connectionString = "Server=$dataSource;Database=$database;Integrated Security=True;"

$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString

$connection.Open()
$query = 'SELECT EName FROM EMP'

$command = $connection.CreateCommand()
$command.CommandText = $query

$result = $command.ExecuteReader()

$JSON = $Result | convertto-json

$JSON