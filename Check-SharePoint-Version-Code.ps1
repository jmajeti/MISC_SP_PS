$connectionString = "Data Source=DBSharePointServer;Integrated Security=SSPI;Initial Catalog=SharePoint_Config_DB"

$cn = new-object system.data.SqlClient.SqlConnection($connectionString)
$ds = new-object "System.Data.DataSet" "dsVersionInfo"
$q = "SELECT * FROM dbo.Versions WHERE VersionId = '00000000-0000-0000-0000-000000000000'"
$da = new-object "System.Data.SqlClient.SqlDataAdapter" ($q, $cn)
$da.Fill($ds)

$dtPerson = new-object "System.Data.DataTable" "dsVersionInfo"
$dtPerson = $ds.Tables[0]
$dtPerson | FOREACH-OBJECT { "VersionID:  " + $_.VersionID + " - Code Version DataBase: " + $_.Version }

"Code Registry: "+ (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\12.0").Version

#RESULT SAMPLE:
# VersionID:  00000000-0000-0000-0000-000000000000 - Code Version DataBase: 12.0.0.6219
# VersionID:  00000000-0000-0000-0000-000000000000 - Code Version DataBase: 12.0.0.6318
# Code Version Registry: 12.0.0.6318

