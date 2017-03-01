# PERI.ListExtender

An extension for List<T> (any List) that extracts data to CSV or PDF

```cs
List<Object>.ExportToCsv(); // Returns StringBuilder
List<Object>.ExportToCsv("[FULL_PATH]"); // Save to file
List<Object>.ExportToPdf("[FULL_PATH]", [PdfFormat]); // Save to file
```

View sample output [here](sample.pdf)

## People to blame

The following personnel are responsible for managing this project.
- [actchua@periapsys.com](mailto:actchua@periapsys.com)

## Developer's Guide

The project uses the ff. technology:
- .Net framework 4
- C# 6.0
- [iTextSharp](https://www.nuget.org/packages/iTextSharp/)

Solution structure:

- PERI.ListExtender
	- The main project
- PERI.ListExtended.Test
	- The Unit Test project
	- Sample codes are available here
	
## Releases

Download release packages [here](https://perilistextender.codeplex.com/releases)
