```


## SyncFusion to Update the word document and convert to pdf.
Try

Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("LicenseKey")
'Loads the template document

Dim document As New WordDocument(in_LetterDocPath, FormatType.Docx)
Console.WriteLine("Balance amount: "+in_BalAmount)
'Find the first occurrence of a particular text in the document

Dim textSelectionContractNo As TextSelection = document.Find("<ContractNumber>", False, True) ' (<ContractNumber>) 
Dim textSelectionEffectiveDate As TextSelection = document.Find("<EffectiveDate>", False, True) ' (<EffectiveDate>) 
Dim textSelectionBalance As TextSelection = document.Find("<Balance>", False, True) ' (<Balance>) 
Dim textSelectionLetterDate As TextSelection = document.Find("<LetterDate>", False, True) ' <Letter Date> 
Dim textSelectionContractNameEn As TextSelection = document.Find("<CompanyNameEn>", False, True) ' (<ContractNumber>) 


Dim textSelectionContractNoAr As TextSelection = document.Find("<ContractNumberAr>", False, True) ' (<ContractNumber>) 
Dim textSelectionEffectiveDateAr As TextSelection = document.Find("<EffectiveDateAr>", False, True) ' (<EffectiveDate>) 
Dim textSelectionBalanceAr As TextSelection = document.Find("<BalanceAr>", False, True) ' (<Balance>) 
Dim textSelectionLetterDateAr As TextSelection = document.Find("<LetterDateAr>", False, True) ' <Letter Date> 
Dim textSelectionContractNameAr As TextSelection = document.Find("<CompanyNameAr>", False, True) ' (<ContractNumber>) 

'Gets the found text as single text range

Dim textRangeContractNo As WTextRange =  textSelectionContractNo.GetAsOneRange()
Dim textRangeEffDate As WTextRange =  textSelectionEffectiveDate.GetAsOneRange()
Dim textRangeBalance As WTextRange =  textSelectionBalance.GetAsOneRange()
Dim textRangeLetterDate As WTextRange =  textSelectionLetterDate.GetAsOneRange()
Dim textRangeContractNameEn As WTextRange =  textSelectionContractNameEn.GetAsOneRange()

Dim textRangeContractNoAr As WTextRange =  textSelectionContractNoAr.GetAsOneRange()
Dim textRangeEffDateAr As WTextRange =  textSelectionEffectiveDateAr.GetAsOneRange()
Dim textRangeBalanceAr As WTextRange =  textSelectionBalanceAr.GetAsOneRange()
Dim textRangeLetterDateAr As WTextRange =  textSelectionLetterDateAr.GetAsOneRange()
Dim textRangeContractNameAr As WTextRange =  textSelectionContractNameAr.GetAsOneRange()


Dim characterformat As WCharacterFormat = New WCharacterFormat(document)
characterformat.TextColor = textRangeBalanceAr.CharacterFormat.TextColor
characterformat.Bold = textRangeBalanceAr.CharacterFormat.Bold
characterformat.FontSize = textRangeBalanceAr.CharacterFormat.FontSize
'characterformat.TextColor = Color.FromArgb(10,0,32,96)

'Modifies the text
textRangeContractNo.Text =in_ContractNumber
textRangeEffDate.Text =in_EffDate
textRangeBalance.Text =in_BalAmount
textRangeLetterDate.Text =Now.Date.ToString("dd/MM/yyyy")
textRangeContractNoAr.Text =in_ContractNumber
textRangeContractNoAr.ApplyCharacterFormat(characterformat)
textRangeEffDateAr.Text =in_EffDate
textRangeEffDateAr.ApplyCharacterFormat(characterformat)
textRangeBalanceAr.Text =in_BalAmount
textRangeBalanceAr.ApplyCharacterFormat(characterformat)
textRangeLetterDateAr.Text =Now.Date.ToString("dd/MM/yyyy")
textRangeLetterDateAr.ApplyCharacterFormat(characterformat)
textRangeContractNameEn.Text =in_CompanyNameEn
textRangeContractNameAr.Text =in_CompanyNameAr


'creates an instance of the DocToPDFConverter - responsible for Word to PDF conversion

Dim converter As New DocToPDFConverter()

'Sets true to enable the fast rendering using direct PDF conversion.
converter.Settings.EnableFastRendering = True

'Converts Word document into PDF document
Dim pdf As PdfDocument = converter.ConvertToPDF(document)

'Saves the PDF file 
pdf.Save(in_LetterPdfPath)

'Closing the PDF And Doc
pdf.Close(True)
document.Close()
Catch Exception As Exception
	console.WriteLine(Exception.Message)
End Try
