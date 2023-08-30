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


