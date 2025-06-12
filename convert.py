from Spire import spire.doc
from Spire import spire.doc.common

# Specify the input and output file paths
inputFile = "index.html"
outputFile = "brandonresume.docx"

# Create an object of the Document class
document = Document()
# Load an HTML file 
document.LoadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.none)

# Save the HTML file to a .docx file
document.SaveToFile(outputFile, FileFormat.Docx2016)
document.Close()