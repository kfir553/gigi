# Create a Word COM object
$word = New-Object -ComObject Word.Application
$word.Visible = $true

# Add a new document
$doc = $word.Documents.Add()
$selection = $word.Selection
$selection.TypeText("Hello from PowerShell!")

# Optional: Save and close
$doc.SaveAs("C:\Temp\Example.docx")
$doc.Close()
$word.Quit()
