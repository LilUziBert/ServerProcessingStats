Import-Module -Name PSSQLite

Add-Type -Path "C:\Program Files\System.Data.SQLite\2010\bin\System.Data.SQLite.dll"
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

$sourceDir = #"\Source Folder Containing RAW Data\*"  ### I prefer not to mess with my original data even if I'm just running processing commands
$destinationDir = #"\Destination Folder Where I Will Be Running My Processing\"

$dateForEmail = Get-Date -UFormat "%m-%d-%Y"  #When sending emails I prefer this format
$dateForFile = Get-Date -UFormat "%Y-%m-%d"   #When storing dates in a DB I prefer this format

$fileName = #"\\ServerStatTimes.xlsx"  #Store one copy on a remote server
$fileNameForEmail = #"C:\ServerStatTimes.xlsx"  #Store a local copy to send out

$xlFixedFormat=[Microsoft.Office.Interop.Excel.XLFileFormat]::xlWorkbookDefault
$xlChart=[Microsoft.Office.Interop.Excel.XLChartType]
$xlOrientation = [Microsoft.Office.Interop.Excel.XlPageOrientation]

$con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
$con.ConnectionString = "Data Source = C:\ServerData.db"
$con.Open()
$dataSource = "C:\ServerData.db"
$sql = $con.CreateCommand()

    #Launch Excel
    $excel = New-Object -ComObject Excel.Application

    #We will make excel invisible
    $excel.visible = $False
    $excel.displayAlerts = $False

    #Open the workbook
    $WB = $excel.Workbooks.Add()

    #I already know how many servers I want to run my test on, so I add a page for each here
    $WB.Worksheets.Add()
    $WB.Worksheets.Add()
    $WB.Worksheets.Add()
    $WB.Worksheets.Add()
    $WB.Worksheets.Add()
    $WB.Worksheets.Add()
    $WB.Worksheets.Add()
    $WB.Worksheets.Add()
    $WB.Worksheets.Add()

function GrabServerData {
    
    $sql.CommandText = "SELECT name FROM server_name"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data = New-Object System.Data.DataSet
    $adapter.Fill($data)

    ForEach ($row in $data.Tables[0].Rows) {
        
        #I have a source folder on each server
        #What this bit of code does:
        # 1) Grabs server name from a SQL DB and stores a UNC path for Source and Destination
        # 2) Checks to see if the Destination exists, and if not creates it and copies all the files from the Source
        # 3) Opens a command prompt and changes directory to the Destination
        # 4) Grabs the current time, and runs my processing files against the Destination (also with the -force argument which my process needs)
        # 5) Grabs the current time again, and subtracts the two to get the elapsed time
        # 6) Writes the server name, date, and seconds to a SQL DB

        $newSource = "\\" + $row.name + $sourceDir
        $newDestination = "\\" + $row.name + $DestinationDir
    
        #Check to see if target folder exists and if it doesnt create it
        if (!(Test-Path -path $newDestination)) {
            New-Item $newDestination -type directory
        }
    
        #Copy files from source folder to target folder
        Copy-Item -path $newSource -destination $newDestination
    
        #Change directory to new folder
        & cd $newDestination
    
        #Get start time
        $startDTM = (Get-Date)
    
        #Point to Serverstat.exe
        $CMD = 'C:\Serverstat.exe'
            
        #Run Serverstat with the force argument
        & $CMD force
    
        #Get end time
        $endDTM = (Get-Date)
    
        #Echo time elapsed
        $result = $(($endDTM-$startDTM).totalseconds)

        $newQuery = "INSERT INTO ServerStatdata (server_name, date, seconds)
                    VALUES (@row, @dateForFile, @result)"
        
        Invoke-SqliteQuery -DataSource $dataSource -Query $newQuery -SqlParameters @{
        
                row = $row.name
                dateForFile = $dateForFile
                result = $result
        }
        $row.name
    }
}

function BuildOverviewWorksheet {
    
    ##############################################################
    #Assign the first worksheet to the variable $serverStatWksht #
    #And then change the name of the worksheet to Overview       #
    ##############################################################

    $serverStatWksht = $excel.Worksheets.Item(1)
    $serverStatWksht.Name = 'Overview'
    $serverStatWksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $serverStatWksht.PageSetup.FitToPagesWide = 1
    $serverStatWksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $serverStatWksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $serverStatWksht.Cells.Item(1,1) = "Date: "
    $serverStatWksht.Cells.Item(1,1).font.bold = $True
    $serverStatWksht.Cells.Item(1,2) = Get-Date -UFormat "%Y-%m-%d"
    $serverStatWksht.Cells.Item(1,2).font.bold = $True

    $serverStatWksht.Cells.Item(4,1) = "Server"
    $serverStatWksht.Cells.Item(4,1).font.bold = $True
    $serverStatWksht.Cells.Item(4,2) = "Today"
    $serverStatWksht.Cells.Item(4,2).font.bold = $True
    $serverStatWksht.Cells.Item(4,3) = "Yesterday"
    $serverStatWksht.Cells.Item(4,3).font.bold = $True
    $serverStatWksht.Cells.Item(4,4) = "Average"
    $serverStatWksht.Cells.Item(4,4).font.bold = $True
    
    #########################################################
    # Fills first column with names of Servers              #
    #########################################################

    $sql = $con.CreateCommand()
    $sql.CommandText = "SELECT name FROM server_name"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data = New-Object System.Data.DataSet
    $adapter.Fill($data)

    $excelRow = 5
    $excelColumn = 1

    foreach ($row in $data.Tables[0].Rows) {
        $serverStatWksht.Cells.Item($excelRow, $excelColumn) = $row.name
        $excelRow++
    }

    $con.Close()

    #########################################################
    # Fills second column with today's times (in seconds)   #
    #########################################################

    $excelRow = 5
    $excelColumn = 2

    $con.Open()
    $sql = $con.CreateCommand()
    $sql.CommandText = "SELECT seconds FROM ServerStatdata WHERE date == date('now')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $dataToday = New-Object System.Data.DataSet
    $adapter.Fill($dataToday)

    foreach ($row in $dataToday.Tables[0].Rows) {
        $serverStatWksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
    }

    $con.Close()

    #########################################################
    # Fills third column with yesterday's times (in seconds)#
    #########################################################

    $excelRow = 5
    $excelColumn = 3

    $con.Open()
    $sql = $con.CreateCommand()
    $sql.CommandText = "SELECT seconds FROM ServerStatdata WHERE date == date('now', '-1 day')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $dataYesterday = New-Object System.Data.DataSet
    $adapter.Fill($dataYesterday)

    foreach ($row in $dataYesterday.Tables[0].Rows) {
        $serverStatWksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
    }

    $con.Close()

    #########################################################
    # Fills fourth column with average time for each server #
    #########################################################

    $excelRow = 5
    $excelColumn = 4

    $con.Open()

    $avgDataSet = New-Object System.Data.DataSet
    $query = "SELECT [server_name], AVG([seconds]) FROM [ServerStatdata] GROUP BY [server_name]"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter($query, $con)
    $adapter.Fill($avgDataSet) | Out-Null
    $con.Close()
    
    $dataAverage = New-Object System.Data.DataTable
    $dataAverage = $avgDataSet.Tables[0] 

    foreach ($row in $dataAverage)  {
        $serverStatWksht.Cells.Item($excelRow, $excelColumn) = $row["AVG([seconds])"].ToString()
        $excelRow++
    }

    $con.Close()

    #########################################################
    # After all the data is pulled, a chart is created      #
    #########################################################

    $dataForChart = $serverStatWksht.Range("Overview!A4:D15").CurrentRegion

    $serverStatWksht.Columns.Item("A:D").EntireColumn.AutoFit() | out-null

    $chart = $serverStatWksht.Shapes.AddChart().Chart
    $chart.ChartType = $xlChart::xlColumnClustered

    $serverStatWksht.shapes.item("Chart 1").top = 25
    $serverStatWksht.shapes.item("Chart 1").left = 250

    $chart.SetSourceData($dataForChart)
    
    #Someone wanted a copy of the chart
    #Here's how to export as a png
    $chart.Export("C:\ServerStatChart.png", "png")

    #########################################################


}

function Server1Worksheet {

    $server1Wksht = $excel.Worksheets.Item(2)
    $server1Wksht.Name = 'Server1'
    $server1Wksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $server1Wksht.PageSetup.FitToPagesWide = 1
    $server1Wksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $server1Wksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $server1Wksht.Cells.Item(1,1) = "Server 1"
    $server1Wksht.Cells.Item(1,1).font.bold = $True
    $server1Wksht.Cells.Item(1,3) = "Date: "
    $server1Wksht.Cells.Item(1,3).font.bold = $True
    $server1Wksht.Cells.Item(1,4) = Get-Date -UFormat "%Y-%m-%d"
    $server1Wksht.Cells.Item(1,4).font.bold = $True
    $server1Wksht.Cells.Item(4,1) = "7 Days of Data"
    $server1Wksht.Cells.Item(4,1).font.bold = $True
    $server1Wksht.Cells.Item(4,4) = "30 Days of Data"
    $server1Wksht.Cells.Item(4,4).font.bold = $True
    

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source = C:\ServerData.db"
    $con.Open()
    $sql = $con.CreateCommand()

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server1' AND date BETWEEN datetime('now','-7 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data7Days = New-Object System.Data.DataSet
    $adapter.Fill($data7Days)

    $excelRow = 6
    $excelColumn = 1

    foreach ($row in $data7Days.Tables[0].Rows) {
        $server1Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server1Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server1' AND date BETWEEN datetime('now','-30 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data30Days = New-Object System.Data.DataSet
    $adapter.Fill($data30Days)

    $excelRow = 6
    $excelColumn = 4

    foreach ($row in $data30Days.Tables[0].Rows) {
        $server1Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server1Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $con.Close()

    $dataFor7DayChart = $server1Wksht.Range("Server1!A6:B12").CurrentRegion
    
    $server1Wksht.Columns.Item("A:F").EntireColumn.AutoFit() | out-null

    $chart7Days = $server1Wksht.Shapes.AddChart().Chart
    $chart7Days.ChartType = $xlChart::xlLine

    $chart7Days.HasTitle = $True
    $chart7Days.ChartTitle.Text = "Last 7 Days of Data"

    $server1Wksht.shapes.item("Chart 1").top = 45
    $server1Wksht.shapes.item("Chart 1").left = 401

    $chart7Days.SetSourceData($dataFor7DayChart)

    $dataFor30DayChart = $server1Wksht.Range("Server1!D6:E34").CurrentRegion

    $chart30Days = $server1Wksht.Shapes.AddChart().Chart
    $chart30Days.ChartType = $xlChart::xlLine

    $chart30Days.HasTitle = $True
    $chart30Days.ChartTitle.Text = "Last 30 Days of Data"

    $server1Wksht.shapes.item("Chart 2").top = 315
    $server1Wksht.shapes.item("Chart 2").left = 401

    $chart30Days.SetSourceData($dataFor30DayChart)

}
    
function Server2Worksheet {

    $server2Wksht = $excel.Worksheets.Item(3)
    $server2Wksht.Name = 'Server2'
    $server2Wksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $server2Wksht.PageSetup.FitToPagesWide = 1
    $server2Wksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $server2Wksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $server2Wksht.Cells.Item(1,1) = "Server 2"
    $server2Wksht.Cells.Item(1,1).font.bold = $True
    $server2Wksht.Cells.Item(1,3) = "Date: "
    $server2Wksht.Cells.Item(1,3).font.bold = $True
    $server2Wksht.Cells.Item(1,4) = Get-Date -UFormat "%Y-%m-%d"
    $server2Wksht.Cells.Item(1,4).font.bold = $True
    $server2Wksht.Cells.Item(4,1) = "7 Days of Data"
    $server2Wksht.Cells.Item(4,1).font.bold = $True
    $server2Wksht.Cells.Item(4,4) = "30 Days of Data"
    $server2Wksht.Cells.Item(4,4).font.bold = $True
    

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source = C:\ServerData.db"
    $con.Open()
    $sql = $con.CreateCommand()

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server2' AND date BETWEEN datetime('now','-7 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data7Days = New-Object System.Data.DataSet
    $adapter.Fill($data7Days)

    $excelRow = 6
    $excelColumn = 1

    foreach ($row in $data7Days.Tables[0].Rows) {
        $server2Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server2Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server2' AND date BETWEEN datetime('now','-30 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data30Days = New-Object System.Data.DataSet
    $adapter.Fill($data30Days)

    $excelRow = 6
    $excelColumn = 4

    foreach ($row in $data30Days.Tables[0].Rows) {
        $server2Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server2Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $con.Close()

    $dataFor7DayChart = $server2Wksht.Range("Server2!A6:B12").CurrentRegion
    
    $server2Wksht.Columns.Item("A:F").EntireColumn.AutoFit() | out-null

    $chart7Days = $server2Wksht.Shapes.AddChart().Chart
    $chart7Days.ChartType = $xlChart::xlLine

    $chart7Days.HasTitle = $True
    $chart7Days.ChartTitle.Text = "Last 7 Days of Data"

    $server2Wksht.shapes.item("Chart 1").top = 45
    $server2Wksht.shapes.item("Chart 1").left = 401

    $chart7Days.SetSourceData($dataFor7DayChart)

    $dataFor30DayChart = $server2Wksht.Range("Server2!D6:E34").CurrentRegion

    $chart30Days = $server2Wksht.Shapes.AddChart().Chart
    $chart30Days.ChartType = $xlChart::xlLine

    $chart30Days.HasTitle = $True
    $chart30Days.ChartTitle.Text = "Last 30 Days of Data"

    $server2Wksht.shapes.item("Chart 2").top = 315
    $server2Wksht.shapes.item("Chart 2").left = 401

    $chart30Days.SetSourceData($dataFor30DayChart)

}

function Server3Worksheet {

    $server3Wksht = $excel.Worksheets.Item(4)
    $server3Wksht.Name = 'Server3'
    $server3Wksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $server3Wksht.PageSetup.FitToPagesWide = 1
    $server3Wksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $server3Wksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $server3Wksht.Cells.Item(1,1) = "Server 3"
    $server3Wksht.Cells.Item(1,1).font.bold = $True
    $server3Wksht.Cells.Item(1,3) = "Date: "
    $server3Wksht.Cells.Item(1,3).font.bold = $True
    $server3Wksht.Cells.Item(1,4) = Get-Date -UFormat "%Y-%m-%d"
    $server3Wksht.Cells.Item(1,4).font.bold = $True
    $server3Wksht.Cells.Item(4,1) = "7 Days of Data"
    $server3Wksht.Cells.Item(4,1).font.bold = $True
    $server3Wksht.Cells.Item(4,4) = "30 Days of Data"
    $server3Wksht.Cells.Item(4,4).font.bold = $True
    

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source = C:\ServerData.db"
    $con.Open()
    $sql = $con.CreateCommand()

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server3' AND date BETWEEN datetime('now','-7 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data7Days = New-Object System.Data.DataSet
    $adapter.Fill($data7Days)

    $excelRow = 6
    $excelColumn = 1

    foreach ($row in $data7Days.Tables[0].Rows) {
        $server3Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server3Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server3' AND date BETWEEN datetime('now','-30 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data30Days = New-Object System.Data.DataSet
    $adapter.Fill($data30Days)

    $excelRow = 6
    $excelColumn = 4

    foreach ($row in $data30Days.Tables[0].Rows) {
        $server3Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server3Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $con.Close()

    $dataFor7DayChart = $server3Wksht.Range("Server3!A6:B12").CurrentRegion
    
    $server3Wksht.Columns.Item("A:F").EntireColumn.AutoFit() | out-null

    $chart7Days = $server3Wksht.Shapes.AddChart().Chart
    $chart7Days.ChartType = $xlChart::xlLine

    $chart7Days.HasTitle = $True
    $chart7Days.ChartTitle.Text = "Last 7 Days of Data"

    $server3Wksht.shapes.item("Chart 1").top = 45
    $server3Wksht.shapes.item("Chart 1").left = 401

    $chart7Days.SetSourceData($dataFor7DayChart)

    $dataFor30DayChart = $server3Wksht.Range("Server!D6:E34").CurrentRegion

    $chart30Days = $server3Wksht.Shapes.AddChart().Chart
    $chart30Days.ChartType = $xlChart::xlLine

    $chart30Days.HasTitle = $True
    $chart30Days.ChartTitle.Text = "Last 30 Days of Data"

    $server3Wksht.shapes.item("Chart 2").top = 315
    $server3Wksht.shapes.item("Chart 2").left = 401

    $chart30Days.SetSourceData($dataFor30DayChart)

}

function Server4Worksheet {

    $server4Wksht = $excel.Worksheets.Item(5)
    $server4Wksht.Name = 'Server4'
    $server4Wksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $server4Wksht.PageSetup.FitToPagesWide = 1
    $server4Wksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $server4Wksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $server4Wksht.Cells.Item(1,1) = "Server 4"
    $server4Wksht.Cells.Item(1,1).font.bold = $True
    $server4Wksht.Cells.Item(1,3) = "Date: "
    $server4Wksht.Cells.Item(1,3).font.bold = $True
    $server4Wksht.Cells.Item(1,4) = Get-Date -UFormat "%Y-%m-%d"
    $server4Wksht.Cells.Item(1,4).font.bold = $True
    $server4Wksht.Cells.Item(4,1) = "7 Days of Data"
    $server4Wksht.Cells.Item(4,1).font.bold = $True
    $server4Wksht.Cells.Item(4,4) = "30 Days of Data"
    $server4Wksht.Cells.Item(4,4).font.bold = $True
    

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source = C:\ServerData.db"
    $con.Open()
    $sql = $con.CreateCommand()

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server4' AND date BETWEEN datetime('now','-7 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data7Days = New-Object System.Data.DataSet
    $adapter.Fill($data7Days)

    $excelRow = 6
    $excelColumn = 1

    foreach ($row in $data7Days.Tables[0].Rows) {
        $server4Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server4Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server4' AND date BETWEEN datetime('now','-30 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data30Days = New-Object System.Data.DataSet
    $adapter.Fill($data30Days)

    $excelRow = 6
    $excelColumn = 4

    foreach ($row in $data30Days.Tables[0].Rows) {
        $server4Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server4Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $con.Close()

    $dataFor7DayChart = $server4Wksht.Range("Server4!A6:B12").CurrentRegion
    
    $server4Wksht.Columns.Item("A:F").EntireColumn.AutoFit() | out-null

    $chart7Days = $server4Wksht.Shapes.AddChart().Chart
    $chart7Days.ChartType = $xlChart::xlLine

    $chart7Days.HasTitle = $True
    $chart7Days.ChartTitle.Text = "Last 7 Days of Data"

    $server4Wksht.shapes.item("Chart 1").top = 45
    $server4Wksht.shapes.item("Chart 1").left = 401

    $chart7Days.SetSourceData($dataFor7DayChart)

    $dataFor30DayChart = $server4Wksht.Range("Server4!D6:E34").CurrentRegion

    $chart30Days = $server4Wksht.Shapes.AddChart().Chart
    $chart30Days.ChartType = $xlChart::xlLine

    $chart30Days.HasTitle = $True
    $chart30Days.ChartTitle.Text = "Last 30 Days of Data"

    $server4Wksht.shapes.item("Chart 2").top = 315
    $server4Wksht.shapes.item("Chart 2").left = 401

    $chart30Days.SetSourceData($dataFor30DayChart)

}

function Server5Worksheet {

    $server5Wksht = $excel.Worksheets.Item(6)
    $server5Wksht.Name = 'Server5'
    $server5Wksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $server5Wksht.PageSetup.FitToPagesWide = 1
    $server5Wksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $server5Wksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $server5Wksht.Cells.Item(1,1) = "Server 5"
    $server5Wksht.Cells.Item(1,1).font.bold = $True
    $server5Wksht.Cells.Item(1,3) = "Date: "
    $server5Wksht.Cells.Item(1,3).font.bold = $True
    $server5Wksht.Cells.Item(1,4) = Get-Date -UFormat "%Y-%m-%d"
    $server5Wksht.Cells.Item(1,4).font.bold = $True
    $server5Wksht.Cells.Item(4,1) = "7 Days of Data"
    $server5Wksht.Cells.Item(4,1).font.bold = $True
    $server5Wksht.Cells.Item(4,4) = "30 Days of Data"
    $server5Wksht.Cells.Item(4,4).font.bold = $True
    

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source = C:\ServerData.db"
    $con.Open()
    $sql = $con.CreateCommand()

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server5' AND date BETWEEN datetime('now','-7 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data7Days = New-Object System.Data.DataSet
    $adapter.Fill($data7Days)

    $excelRow = 6
    $excelColumn = 1

    foreach ($row in $data7Days.Tables[0].Rows) {
        $server5Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server5Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server5' AND date BETWEEN datetime('now','-30 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data30Days = New-Object System.Data.DataSet
    $adapter.Fill($data30Days)

    $excelRow = 6
    $excelColumn = 4

    foreach ($row in $data30Days.Tables[0].Rows) {
        $server5Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server5Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $con.Close()

    $dataFor7DayChart = $server5Wksht.Range("Server5!A6:B12").CurrentRegion
    
    $server5Wksht.Columns.Item("A:F").EntireColumn.AutoFit() | out-null

    $chart7Days = $server5Wksht.Shapes.AddChart().Chart
    $chart7Days.ChartType = $xlChart::xlLine

    $chart7Days.HasTitle = $True
    $chart7Days.ChartTitle.Text = "Last 7 Days of Data"

    $server5Wksht.shapes.item("Chart 1").top = 45
    $server5Wksht.shapes.item("Chart 1").left = 401

    $chart7Days.SetSourceData($dataFor7DayChart)

    $dataFor30DayChart = $server5Wksht.Range("Server5!D6:E34").CurrentRegion

    $chart30Days = $server5Wksht.Shapes.AddChart().Chart
    $chart30Days.ChartType = $xlChart::xlLine

    $chart30Days.HasTitle = $True
    $chart30Days.ChartTitle.Text = "Last 30 Days of Data"

    $server5Wksht.shapes.item("Chart 2").top = 315
    $server5Wksht.shapes.item("Chart 2").left = 401

    $chart30Days.SetSourceData($dataFor30DayChart)

}

function Server6Worksheet {

    $server6Wksht = $excel.Worksheets.Item(7)
    $server6Wksht.Name = 'Server6'
    $server6Wksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $server6Wksht.PageSetup.FitToPagesWide = 1
    $server6Wksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $server6Wksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $server6Wksht.Cells.Item(1,1) = "Server 6"
    $server6Wksht.Cells.Item(1,1).font.bold = $True
    $server6Wksht.Cells.Item(1,3) = "Date: "
    $server6Wksht.Cells.Item(1,3).font.bold = $True
    $server6Wksht.Cells.Item(1,4) = Get-Date -UFormat "%Y-%m-%d"
    $server6Wksht.Cells.Item(1,4).font.bold = $True
    $server6Wksht.Cells.Item(4,1) = "7 Days of Data"
    $server6Wksht.Cells.Item(4,1).font.bold = $True
    $server6Wksht.Cells.Item(4,4) = "30 Days of Data"
    $server6Wksht.Cells.Item(4,4).font.bold = $True
    

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source = C:\ServerData.db"
    $con.Open()
    $sql = $con.CreateCommand()

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server6' AND date BETWEEN datetime('now','-7 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data7Days = New-Object System.Data.DataSet
    $adapter.Fill($data7Days)

    $excelRow = 6
    $excelColumn = 1

    foreach ($row in $data7Days.Tables[0].Rows) {
        $server6Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server6Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server6' AND date BETWEEN datetime('now','-30 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data30Days = New-Object System.Data.DataSet
    $adapter.Fill($data30Days)

    $excelRow = 6
    $excelColumn = 4

    foreach ($row in $data30Days.Tables[0].Rows) {
        $server6Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server6Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $con.Close()

    $dataFor7DayChart = $server6Wksht.Range("Server6!A6:B12").CurrentRegion
    
    $server6Wksht.Columns.Item("A:F").EntireColumn.AutoFit() | out-null

    $chart7Days = $server6Wksht.Shapes.AddChart().Chart
    $chart7Days.ChartType = $xlChart::xlLine

    $chart7Days.HasTitle = $True
    $chart7Days.ChartTitle.Text = "Last 7 Days of Data"

    $server6Wksht.shapes.item("Chart 1").top = 45
    $server6Wksht.shapes.item("Chart 1").left = 401

    $chart7Days.SetSourceData($dataFor7DayChart)

    $dataFor30DayChart = $server6Wksht.Range("Server6!D6:E34").CurrentRegion

    $chart30Days = $server6Wksht.Shapes.AddChart().Chart
    $chart30Days.ChartType = $xlChart::xlLine

    $chart30Days.HasTitle = $True
    $chart30Days.ChartTitle.Text = "Last 30 Days of Data"

    $server6Wksht.shapes.item("Chart 2").top = 315
    $server6Wksht.shapes.item("Chart 2").left = 401

    $chart30Days.SetSourceData($dataFor30DayChart)

}

function Server7Worksheet {

    $server7Wksht = $excel.Worksheets.Item(8)
    $server7Wksht.Name = 'Server7'
    $server7Wksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $server7Wksht.PageSetup.FitToPagesWide = 1
    $server7Wksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $server7Wksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $server7Wksht.Cells.Item(1,1) = "Server 7"
    $server7Wksht.Cells.Item(1,1).font.bold = $True
    $server7Wksht.Cells.Item(1,3) = "Date: "
    $server7Wksht.Cells.Item(1,3).font.bold = $True
    $server7Wksht.Cells.Item(1,4) = Get-Date -UFormat "%Y-%m-%d"
    $server7Wksht.Cells.Item(1,4).font.bold = $True
    $server7Wksht.Cells.Item(4,1) = "7 Days of Data"
    $server7Wksht.Cells.Item(4,1).font.bold = $True
    $server7Wksht.Cells.Item(4,4) = "30 Days of Data"
    $server7Wksht.Cells.Item(4,4).font.bold = $True
    

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source = C:\ServerData.db"
    $con.Open()
    $sql = $con.CreateCommand()

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server7' AND date BETWEEN datetime('now','-7 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data7Days = New-Object System.Data.DataSet
    $adapter.Fill($data7Days)

    $excelRow = 6
    $excelColumn = 1

    foreach ($row in $data7Days.Tables[0].Rows) {
        $server7Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server7Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server7' AND date BETWEEN datetime('now','-30 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data30Days = New-Object System.Data.DataSet
    $adapter.Fill($data30Days)

    $excelRow = 6
    $excelColumn = 4

    foreach ($row in $data30Days.Tables[0].Rows) {
        $server7Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server7Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $con.Close()

    $dataFor7DayChart = $server7Wksht.Range("Server7!A6:B12").CurrentRegion
    
    $server7Wksht.Columns.Item("A:F").EntireColumn.AutoFit() | out-null

    $chart7Days = $server7Wksht.Shapes.AddChart().Chart
    $chart7Days.ChartType = $xlChart::xlLine

    $chart7Days.HasTitle = $True
    $chart7Days.ChartTitle.Text = "Last 7 Days of Data"

    $server7Wksht.shapes.item("Chart 1").top = 45
    $server7Wksht.shapes.item("Chart 1").left = 401

    $chart7Days.SetSourceData($dataFor7DayChart)

    $dataFor30DayChart = $server7Wksht.Range("Server7!D6:E34").CurrentRegion

    $chart30Days = $server7Wksht.Shapes.AddChart().Chart
    $chart30Days.ChartType = $xlChart::xlLine

    $chart30Days.HasTitle = $True
    $chart30Days.ChartTitle.Text = "Last 30 Days of Data"

    $server7Wksht.shapes.item("Chart 2").top = 315
    $server7Wksht.shapes.item("Chart 2").left = 401

    $chart30Days.SetSourceData($dataFor30DayChart)

}

function Server8Worksheet {

    $server8Wksht = $excel.Worksheets.Item(9)
    $server8Wksht.Name = 'Server8'
    $server8Wksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $server8Wksht.PageSetup.FitToPagesWide = 1
    $server8Wksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $server8Wksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $server8Wksht.Cells.Item(1,1) = "Server 8"
    $server8Wksht.Cells.Item(1,1).font.bold = $True
    $server8Wksht.Cells.Item(1,3) = "Date: "
    $server8Wksht.Cells.Item(1,3).font.bold = $True
    $server8Wksht.Cells.Item(1,4) = Get-Date -UFormat "%Y-%m-%d"
    $server8Wksht.Cells.Item(1,4).font.bold = $True
    $server8Wksht.Cells.Item(4,1) = "7 Days of Data"
    $server8Wksht.Cells.Item(4,1).font.bold = $True
    $server8Wksht.Cells.Item(4,4) = "30 Days of Data"
    $server8Wksht.Cells.Item(4,4).font.bold = $True
    

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source = C:\ServerData.db"
    $con.Open()
    $sql = $con.CreateCommand()

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server8' AND date BETWEEN datetime('now','-7 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data7Days = New-Object System.Data.DataSet
    $adapter.Fill($data7Days)

    $excelRow = 6
    $excelColumn = 1

    foreach ($row in $data7Days.Tables[0].Rows) {
        $server8Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server8Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server8' AND date BETWEEN datetime('now','-30 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data30Days = New-Object System.Data.DataSet
    $adapter.Fill($data30Days)

    $excelRow = 6
    $excelColumn = 4

    foreach ($row in $data30Days.Tables[0].Rows) {
        $server8Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server8Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $con.Close()

    $dataFor7DayChart = $server8Wksht.Range("Server8!A6:B12").CurrentRegion
    
    $server8Wksht.Columns.Item("A:F").EntireColumn.AutoFit() | out-null

    $chart7Days = $server8Wksht.Shapes.AddChart().Chart
    $chart7Days.ChartType = $xlChart::xlLine

    $chart7Days.HasTitle = $True
    $chart7Days.ChartTitle.Text = "Last 7 Days of Data"

    $server8Wksht.shapes.item("Chart 1").top = 45
    $server8Wksht.shapes.item("Chart 1").left = 401

    $chart7Days.SetSourceData($dataFor7DayChart)

    $dataFor30DayChart = $server8Wksht.Range("Server8!D6:E34").CurrentRegion

    $chart30Days = $server8Wksht.Shapes.AddChart().Chart
    $chart30Days.ChartType = $xlChart::xlLine

    $chart30Days.HasTitle = $True
    $chart30Days.ChartTitle.Text = "Last 30 Days of Data"

    $server8Wksht.shapes.item("Chart 2").top = 315
    $server8Wksht.shapes.item("Chart 2").left = 401

    $chart30Days.SetSourceData($dataFor30DayChart)

}

function Server9Worksheet {

    $server9Wksht = $excel.Worksheets.Item(10)
    $server9Wksht.Name = 'Server9'
    $server9Wksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $server9Wksht.PageSetup.FitToPagesWide = 1
    $server9Wksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $server9Wksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $server9Wksht.Cells.Item(1,1) = "Server 9"
    $server9Wksht.Cells.Item(1,1).font.bold = $True
    $server9Wksht.Cells.Item(1,3) = "Date: "
    $server9Wksht.Cells.Item(1,3).font.bold = $True
    $server9Wksht.Cells.Item(1,4) = Get-Date -UFormat "%Y-%m-%d"
    $server9Wksht.Cells.Item(1,4).font.bold = $True
    $server9Wksht.Cells.Item(4,1) = "7 Days of Data"
    $server9Wksht.Cells.Item(4,1).font.bold = $True
    $server9Wksht.Cells.Item(4,4) = "30 Days of Data"
    $server9Wksht.Cells.Item(4,4).font.bold = $True
    

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source = C:\ServerData.db"
    $con.Open()
    $sql = $con.CreateCommand()

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server9' AND date BETWEEN datetime('now','-7 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data7Days = New-Object System.Data.DataSet
    $adapter.Fill($data7Days)

    $excelRow = 6
    $excelColumn = 1

    foreach ($row in $data7Days.Tables[0].Rows) {
        $server9Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server9Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server9' AND date BETWEEN datetime('now','-30 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data30Days = New-Object System.Data.DataSet
    $adapter.Fill($data30Days)

    $excelRow = 6
    $excelColumn = 4

    foreach ($row in $data30Days.Tables[0].Rows) {
        $server9Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server9Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $con.Close()

    $dataFor7DayChart = $server9Wksht.Range("Server9!A6:B12").CurrentRegion
    
    $server9Wksht.Columns.Item("A:F").EntireColumn.AutoFit() | out-null

    $chart7Days = $server9Wksht.Shapes.AddChart().Chart
    $chart7Days.ChartType = $xlChart::xlLine

    $chart7Days.HasTitle = $True
    $chart7Days.ChartTitle.Text = "Last 7 Days of Data"

    $server9Wksht.shapes.item("Chart 1").top = 45
    $server9Wksht.shapes.item("Chart 1").left = 401

    $chart7Days.SetSourceData($dataFor7DayChart)

    $dataFor30DayChart = $server9Wksht.Range("Server9!D6:E34").CurrentRegion

    $chart30Days = $server9Wksht.Shapes.AddChart().Chart
    $chart30Days.ChartType = $xlChart::xlLine

    $chart30Days.HasTitle = $True
    $chart30Days.ChartTitle.Text = "Last 30 Days of Data"

    $server9Wksht.shapes.item("Chart 2").top = 315
    $server9Wksht.shapes.item("Chart 2").left = 401

    $chart30Days.SetSourceData($dataFor30DayChart)

}

function Server10Worksheet {

    $server10Wksht = $excel.Worksheets.Item(11)
    $server10Wksht.Name = 'Server10'
    $server10Wksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $server10Wksht.PageSetup.FitToPagesWide = 1
    $server10Wksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $server10Wksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $server10Wksht.Cells.Item(1,1) = "Server 10"
    $server10Wksht.Cells.Item(1,1).font.bold = $True
    $server10Wksht.Cells.Item(1,3) = "Date: "
    $server10Wksht.Cells.Item(1,3).font.bold = $True
    $server10Wksht.Cells.Item(1,4) = Get-Date -UFormat "%Y-%m-%d"
    $server10Wksht.Cells.Item(1,4).font.bold = $True
    $server10Wksht.Cells.Item(4,1) = "7 Days of Data"
    $server10Wksht.Cells.Item(4,1).font.bold = $True
    $server10Wksht.Cells.Item(4,4) = "30 Days of Data"
    $server10Wksht.Cells.Item(4,4).font.bold = $True
    

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source = C:\ServerData.db"
    $con.Open()
    $sql = $con.CreateCommand()

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server10' AND date BETWEEN datetime('now','-7 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data7Days = New-Object System.Data.DataSet
    $adapter.Fill($data7Days)

    $excelRow = 6
    $excelColumn = 1

    foreach ($row in $data7Days.Tables[0].Rows) {
        $server10Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server10Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server10' AND date BETWEEN datetime('now','-30 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data30Days = New-Object System.Data.DataSet
    $adapter.Fill($data30Days)

    $excelRow = 6
    $excelColumn = 4

    foreach ($row in $data30Days.Tables[0].Rows) {
        $server10Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server10Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $con.Close()

    $dataFor7DayChart = $server10Wksht.Range("Server10!A6:B12").CurrentRegion
    
    $server10Wksht.Columns.Item("A:F").EntireColumn.AutoFit() | out-null

    $chart7Days = $server10Wksht.Shapes.AddChart().Chart
    $chart7Days.ChartType = $xlChart::xlLine

    $chart7Days.HasTitle = $True
    $chart7Days.ChartTitle.Text = "Last 7 Days of Data"

    $server10Wksht.shapes.item("Chart 1").top = 45
    $server10Wksht.shapes.item("Chart 1").left = 401

    $chart7Days.SetSourceData($dataFor7DayChart)

    $dataFor30DayChart = $server10Wksht.Range("Server10!D6:E34").CurrentRegion

    $chart30Days = $server10Wksht.Shapes.AddChart().Chart
    $chart30Days.ChartType = $xlChart::xlLine

    $chart30Days.HasTitle = $True
    $chart30Days.ChartTitle.Text = "Last 30 Days of Data"

    $server10Wksht.shapes.item("Chart 2").top = 315
    $server10Wksht.shapes.item("Chart 2").left = 401

    $chart30Days.SetSourceData($dataFor30DayChart)

}

function Server11Worksheet {

    $server11Wksht = $excel.Worksheets.Item(12)
    $server11Wksht.Name = 'Server11'
    $server11Wksht.PageSetup.Orientation = $xlOrientation::xlLandscape
    $server11Wksht.PageSetup.FitToPagesWide = 1
    $server11Wksht.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $server11Wksht.PageSetup.RightMargin = $excel.InchesToPoints(0.25)

    $server11Wksht.Cells.Item(1,1) = "Server 11"
    $server11Wksht.Cells.Item(1,1).font.bold = $True
    $server11Wksht.Cells.Item(1,3) = "Date: "
    $server11Wksht.Cells.Item(1,3).font.bold = $True
    $server11Wksht.Cells.Item(1,4) = Get-Date -UFormat "%Y-%m-%d"
    $server11Wksht.Cells.Item(1,4).font.bold = $True
    $server11Wksht.Cells.Item(4,1) = "7 Days of Data"
    $server11Wksht.Cells.Item(4,1).font.bold = $True
    $server11Wksht.Cells.Item(4,4) = "30 Days of Data"
    $server11Wksht.Cells.Item(4,4).font.bold = $True
    

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source = C:\ServerData.db"
    $con.Open()
    $sql = $con.CreateCommand()

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server11' AND date BETWEEN datetime('now','-7 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data7Days = New-Object System.Data.DataSet
    $adapter.Fill($data7Days)

    $excelRow = 6
    $excelColumn = 1

    foreach ($row in $data7Days.Tables[0].Rows) {
        $server11Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server11Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $sql.CommandText = "SELECT [seconds], [date] FROM [ServerStatdata] WHERE [server_name] == 'Server11' AND date BETWEEN datetime('now','-30 days') AND datetime('now','localtime')"
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data30Days = New-Object System.Data.DataSet
    $adapter.Fill($data30Days)

    $excelRow = 6
    $excelColumn = 4

    foreach ($row in $data30Days.Tables[0].Rows) {
        $server11Wksht.Cells.Item($excelRow, $excelColumn) = $row.date
        $excelColumn++
        $server11Wksht.Cells.Item($excelRow, $excelColumn) = $row.seconds
        $excelRow++
        $excelColumn--
    }

    $con.Close()

    $dataFor7DayChart = $server11Wksht.Range("Server11!A6:B12").CurrentRegion
    
    $server11Wksht.Columns.Item("A:F").EntireColumn.AutoFit() | out-null

    $chart7Days = $server11Wksht.Shapes.AddChart().Chart
    $chart7Days.ChartType = $xlChart::xlLine

    $chart7Days.HasTitle = $True
    $chart7Days.ChartTitle.Text = "Last 7 Days of Data"

    $server11Wksht.shapes.item("Chart 1").top = 45
    $server11Wksht.shapes.item("Chart 1").left = 401

    $chart7Days.SetSourceData($dataFor7DayChart)

    $dataFor30DayChart = $server11Wksht.Range("Server11!D6:E34").CurrentRegion

    $chart30Days = $server11Wksht.Shapes.AddChart().Chart
    $chart30Days.ChartType = $xlChart::xlLine

    $chart30Days.HasTitle = $True
    $chart30Days.ChartTitle.Text = "Last 30 Days of Data"

    $server11Wksht.shapes.item("Chart 2").top = 315
    $server11Wksht.shapes.item("Chart 2").left = 401

    $chart30Days.SetSourceData($dataFor30DayChart)

}

function SendOutEmails {
    
    $emailSmtpServer = ""   #SMTP server goes here
    $emailSmtpServerPort = "25"


    $emailMessage = New-Object System.Net.Mail.MailMessage
    $emailMessage.From = "Vibert <vibert@TestEmail.com>"
    $emailMessage.To.Add( "Vibert <vibert@TestEmail.com>" )
    $emailMessage.Subject = "ServerStat Report for $dateForEmail"
    $emailMessage.IsBodyHtml = $true
    
    $emailMessage.Body = 
    "<p>Hello,</p>
    <p>This email contains the daily ServerStat Report for $dateForEmail! </p>"

    $SMTPClient = New-Object System.Net.Mail.SmtpClient ( $emailSmtpServer , $emailSmtpServerPort)

    $attachment = "C:\ServerStatTimes.xlsx"
    $emailMessage.Attachments.Add( $attachment )

    $SMTPClient.Send( $emailMessage )
}

GrabTurboData
BuildOverviewWorksheet
Server1Worksheet
Server2Worksheet
Server3Worksheet
Server4Worksheet
Server5Worksheet
Server6Worksheet
Server7Worksheet
Server8Worksheet
Server9Worksheet
Server10Worksheet
Server11Worksheet

    #Save as Excel workbook
    $excel.ActiveWorkbook.SaveAs($fileName, $xlFixedFormat)
    $excel.ActiveWorkbook.SaveAs($fileNameForEmail, $xlFixedFormat)

    $WB.Close()
    $excel.Quit()

    while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)){}
    while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($WB)){}

SendOutEmails
