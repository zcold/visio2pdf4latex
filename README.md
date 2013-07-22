visio2pdf4latex
===============

visio2pdf4latex = Visio to PDF for LaTeX

_**Motivation**_

Before:

1. Draw a figure in visio.
2. **Manually** Save it as an emf file and **font information is LOST**.
3. **Manually** use the [emf2eps tool from LyX] to convert it into eps.
4. Include this _just fine_ figure in LaTeX.

After:

1. Draw a figure in visio.
2. **Automatically** invoke visio2pdf4latex.
3. Include the produced _perfect_ pdf in LaTeX.

_**Usage**_

1. Open visio2pdf4latex.sln in Visual Studio 2012.
2. Build solution.
3. Move visio2pdf4latex.exe and Microsoft.Office.Interop.Visio.dll to the folder you want.
4. Execute visio2pdf4latex example.vsdx

_**Functionality**_

1. Open the visio file ("Example.vsdx")in Visio 2013 in [invsible mode].
2. Export the first shape in the first page to "temp.svg".
3. Open "temp.svg" in Visio 2013.
4. Export the "temp.svg" to "Example.pdf"
5. Quit Visio 2013

**BE SURE** you have at least 1mm margin to the bottom.

Example:

1. Make a rectanle shape as the background as follows.

![Example](https://copy.com/ZhOqkZI0p7Ym/example.png?revision=273, "example")

2. Remove its line color
3. Group all shapes
4. Use visio2pdf4latex

_**Requirement**_

1. [MS .NET Framework 4.5]
2. Visio (only Visio 2013 is tested)
3. Visual Studio 2012

[emf2eps tool from LyX]: http://wiki.lyx.org/Windows/MetafileToEPSConverter
[MS .NET Framework 4.5]: http://www.microsoft.com/en-us/download/details.aspx?id=30653
[invsible mode]: http://msdn.microsoft.com/en-us/library/ff766890.aspx
