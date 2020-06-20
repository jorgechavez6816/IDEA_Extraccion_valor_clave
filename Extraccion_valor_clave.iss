Sub Main
	Call KeyValueExtraction()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Datos: Extracción por valor clave
Function KeyValueExtraction
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.KeyValueExtraction
	dim myArray(5,0)
	myArray(0,0) = "01"
	myArray(1,0) = "02"
	myArray(2,0) = "03"
	myArray(3,0) = "04"
	myArray(4,0) = "05"
	myArray(5,0) = "06"
	task.IncludeAllFields
	task.AddKey "COD_PROD", "A"
	task.DBPrefix = "PROD"
	
	task.CreateMultipleDatabases = TRUE
	task.CreateVirtualDatabase = False
	
	task.ValuesToExtract myArray
	task.PerformTask
	dbName = task.DBName
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase(dbName)
End Function



	
