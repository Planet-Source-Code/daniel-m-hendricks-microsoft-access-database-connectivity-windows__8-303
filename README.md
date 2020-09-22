<div align="center">

## Microsoft Access Database Connectivity \(Windows\)


</div>

### Description

This is an example of how you can pull data from a Microsoft Access database through ADO. Windows only.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel M\. Hendricks](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-m-hendricks.md)
**Level**          |Intermediate
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |PHP 4\.0
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__8-5.md)
**World**          |[PHP](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/php.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-m-hendricks-microsoft-access-database-connectivity-windows__8-303/archive/master.zip)





### Source Code

```
<?
$conn = new COM("ADODB.Connection") or die("Cannot start ADO");
// Microsoft Access connection string.
$conn->Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=C:\\inetpub\\wwwroot\\php\\mydb.mdb");
// SQL statement to build recordset.
$rs = $conn->Execute("SELECT myfield FROM mytable");
echo "<p>Below is a list of values in the MYDB.MDB database, MYABLE table, MYFIELD field.</p>";
// Display all the values in the recordset
while (!$rs->EOF) {
$fv = $rs->Fields("myfield");
echo "Value: ".$fv->value."<br>\n";
$rs->MoveNext();
}
$rs->Close();
?>
```

