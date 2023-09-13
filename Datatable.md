# Code Snippet for Quick Dev

## Convert HTML to Dataset using vb.net

```vb.net
Try
	'Initailize a html document
	Dim doc As New HtmlDocument()
	'Initailize the DataSet
	Dim Ds As New DataSet
	'load the html
	doc.LoadHtml(html)
	'load all the tables
	Dim tables As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//table")

For Each table As HtmlNode In tables
	'Initailize the Dt
	Dim Dt As New DataTable()
	'Load the tr tags of table
	Dim rows As HtmlNodeCollection = table.SelectNodes(".//tr")
	If rows IsNot Nothing Then
		Dim maxTdCount As Integer = 0
		'Check the maximum  td tags exists to create a columns
			For Each trNode  In rows
			' Count the number of <td> elements within the <tr>
					Dim tdCount As Integer = trNode.SelectNodes(".//td").Count
					' Update the maxTdCount if necessary
					If tdCount > maxTdCount Then
						maxTdCount = tdCount
					End If
			Next
    		'Add Columns
        	For Each Value In  Enumerable.Range(0, maxTdCount)
        			Dt.Columns.Add("Column"+Value.ToString,  Type.GetType("System.String"))
        	
        	Next		
	
	
        For Each row In rows
            'Initailize datarow and Index
        	Dim Index As Int32 = 0
        	Dim dr As DataRow = Dt.NewRow()
        	'Load all the td tags of trNode
        	Dim cells As HtmlNodeCollection = row.SelectNodes("td")
        	If cells IsNot Nothing Then
            	For Each cell In cells
                    'If inner text has garbage values  replace it with respect to the values.
	            Dim cellvalue = cell.InnerText.Replace("�","").Replace("&nbsp;","").Replace("&amp;","&")
                    dr("Column"+ index.ToString) = cellvalue
            	    Index = Index+1
              	Next
        	    Dt.rows.Add(dr)
            End If
        Next
        'Add the dt to Dataset
        If  Dt IsNot Nothing  andalso Dt.rows.Count >0 Then
        	Ds.Tables.Add(Dt)
        	Console.WriteLine("Added to Dataset")
        End If
    End If
Next
'Pass the Dataset to outside
io_Ds = Ds
Catch exp As Exception
	Console.WriteLine(exp.Message)
End Try

```


## VlookUp and update the Current month file with respect to the Previous Month.
```vb.net
Try
'Set License Key
SpreadsheetInfo.SetLicense(in_Config("GemboxLicenseKey").ToString)
'Load the Excel File
Dim workbook  As ExcelFile = ExcelFile.Load(in_CurrentMonthPath)
'Laod the required SheetName
Dim worksheet As ExcelWorksheet = workbook.Worksheets(in_Config("UpdatedSheetName").ToString)
'Take the count of rows in the sheet
Dim LastRow = worksheet.Rows.Count
'Initialize the required variables
Dim MaxCellIndex As Integer
Dim FirstRow As Integer = 2
'For each specified columns update the formula in the sheet for the particular columns.
For Each ColumnName As String In in_Config("SpecifiedColumn").ToString.Split(","c).ToList
	'Assign vlookup formula for the column
	Dim VlookupFormula = "="+String.Format( in_Config( ColumnName).ToString, path.GetDirectoryName( in_PreviousMonthPath), path.GetFileName( in_PreviousMonthPath),in_Config("FormattedSheetName").ToString)
	'Assign first cell index to write the formula
	Dim ColumnIndex = Chr(Asc(System.Text.RegularExpressions.Regex.Match( VlookupFormula,"(\w)(?=\=)").ToString)+1).ToString
	'Assign Lastrow as temp
	Dim TempRow = LastRow
	Dim AllRowUpdated  As Integer
	Do 
		'Check row is more than 5000 because gembox as only maxlimit to write cell with 5000
		If TempRow >5000
			 MaxCellIndex = 5000
			 worksheet.Cells.GetSubrange(ColumnIndex+FirstRow.ToString+":"+ColumnIndex+MaxCellIndex.ToString).Formula = VlookupFormula
			TempRow = TempRow - MaxCellIndex
			FirstRow = MaxCellIndex+1
			AllRowUpdated = TempRow
		 Else
		worksheet.Cells.GetSubrange(ColumnIndex+FirstRow.ToString+":"+ColumnIndex+LastRow.ToString).Formula = VlookupFormula
		AllRowUpdated = TempRow-(LastRow-FirstRow+2)
		FirstRow = 2
	End If

	Loop While AllRowUpdated>0
	
Next
'Once wrote formula apply the calculate to do validation.
worksheet.Calculate()

workbook.Save(in_CurrentMonthPath)
console.WriteLine("Added Vlookup formula successfully.")
   Catch ex As Exception
	   Console.WriteLine(ex.Message)
   End Try
        
       
