visio2pdf4latex
===============

visio2pdf4latex = Visio to PDF for LaTeX

**Usage**

1. Open visio2pdf4latex.sln in Visual Studio 2012.
2. Build solution.
3. Move visio2pdf4latex.exe and Microsoft.Office.Interop.Visio.dll to the folder you want.
4. Execute visio2pdf4latex example.vsdx

**Functionality**

1. Open the visio file ("Example.vsdx")in Visio 2013 in [invsible mode].
1. Export the first shape in the first page to "temp.svg".
2. Open "temp.svg" in Visio 2013.
3. Export the "temp.svg" to "Example.pdf"
4. Quit Visio 2013

**BE SURE** you have at least 1mm margin to the bottom.

[invsible mode]: http://msdn.microsoft.com/en-us/library/ff766890.aspx
