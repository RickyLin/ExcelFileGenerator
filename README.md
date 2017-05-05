# ExcelFileGenerator
A lightweight .NET Standard 1.3 library to generate Excel files (.xlsx) using OpenXML

## Features
- Generate Excel files based on a collection of objects
- Retrieve information from attributes
- Support predefined and custom number format

## Sample
1. Define an entity
``` C#
using RL.ExcelFileGenerator;

// ...

public class Product
{
	[ExcelColumn("Product Id", 4)]
	public int Id { get; set; }

	public string Code { get; set; }

	[ExcelColumn("Product Name", 1)]
	public string Name { get; set; }

	[ExcelColumn("Product Price", 3, NumberingFormatString = "#,##0.00")]
	public decimal? Price { get; set; }

	[ExcelColumn("Product Discount", 2)]
	public float Discount { get; set; }

	[ExcelColumn("Available Date", 5, NumberFormatId = 22)]
	public DateTime AvailableDate { get; set; }

	[ExcelColumn("Is Offline", 6)]
	public bool IsOffline { get; set; }
}
```
2. Invoke the Excel static methods to generate .xlsx files.
``` C#
using RL.ExcelFileGenerator;
using System.IO;

// ...

class Program
{
	static void Main(string[] args)
	{
		List<Product> products = new List<Product>(3);
		products.Add(new Product()
			{
				Code = "P_A",
				Discount = 0.1F,
				Id = 1,
				Name = "ABC",
				Price = 6000.85m,
				AvailableDate = DateTime.Now
			});
		products.Add(new Product()
		{
			Code = "P_B",
			Discount = 0.2F,
			Id = 2,
			Name = "XYZ",
			Price = 5500.50m,
			AvailableDate = DateTime.Now.AddDays(-3),
			IsOffline = true
		});
		products.Add(new Product()
			{
				Code = "P_C",
				Discount = 0.22F,
				Id = 3,
				Name = "123",
				Price = null,
				AvailableDate = DateTime.Now.AddDays(-3),
				IsOffline = true,
				NullableBool = true
			});
		products.Add(new Product()
		{
			Code = "P_D",
			Discount = 0.35F,
			Id = 4,
			Name = "456",
			Price = 5800.0m,
			AvailableDate = DateTime.Now.AddDays(-5)
		});

		string fileName = Path.Combine(Directory.GetCurrentDirectory(), "Test.xlsx");

		// create .xlsx file as MemoryStream
		using (MemoryStream ms = ExcelGenerator.CreateStream(products, "Products"))
		{
			using (FileStream fs = File.Create(fileName))
			{
				ms.WriteTo(fs);
			}
		}

		// create .xlsx file
		//ExcelGenerator.CreateFile(products, "Products", fileName);
	}
}

```

## Notes
- the Index property of ExcelColumnAttribute is used in sorting columns, it's not the column index. so for 3 properties of entity have 1, 2, 5, the third property won't be in the column "E" but the column "C".

~~## Nuget Package
`PM> Install-Package RL.ExcelFileGenerator`~~

## FYI
### Predefined number format in Excel
_The following information comes from [ClosedXML](https://closedxml.codeplex.com/wikipage?title=NumberFormatId%20Lookup%20Table&referringTitle=Styles%20-%20NumberFormat)_

Id	|	Format Code
---	|	---
0	|	General
1	|	0
2	|	0.00
3	|	#,##0
4	|	#,##0.00
9	|	0%
10	|	0.00%
11	|	0.00E+00
12	|	# ?/?
13	|	# ??/??
14	|	d/m/yyyy
15	|	d-mmm-yy
16	|	d-mmm
17	|	mmm-yy
18	|	h:mm tt
19	|	h:mm:ss tt
20	|	H:mm
21	|	H:mm:ss
22	|	m/d/yyyy H:mm
37	|	#,##0 ;(#,##0)
38	|	#,##0 ;\[Red\](#,##0)
39	|	#,##0.00;(#,##0.00)
40	|	#,##0.00;\[Red\](#,##0.00)
45	|	mm:ss
46	|	[h]:mm:ss
47	|	mmss.0
48	|	##0.0E+0
49	|	@

### MIME for MS Office Files
Ext	|	MIME
---	|	---
.doc	|	application/msword
.dot	|	application/msword
.docx	|	application/vnd.openxmlformats-officedocument.wordprocessingml.document
.dotx	|	application/vnd.openxmlformats-officedocument.wordprocessingml.template
.docm	|	application/vnd.ms-word.document.macroEnabled.12
.dotm	|	application/vnd.ms-word.template.macroEnabled.12
.xls	|	application/vnd.ms-excel
.xlt	|	application/vnd.ms-excel
.xla	|	application/vnd.ms-excel
.xlsx	|	application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
.xltx	|	application/vnd.openxmlformats-officedocument.spreadsheetml.template
.xlsm	|	application/vnd.ms-excel.sheet.macroEnabled.12
.xltm	|	application/vnd.ms-excel.template.macroEnabled.12
.xlam	|	application/vnd.ms-excel.addin.macroEnabled.12
.xlsb	|	application/vnd.ms-excel.sheet.binary.macroEnabled.12
.ppt	|	application/vnd.ms-powerpoint
.pot	|	application/vnd.ms-powerpoint
.pps	|	application/vnd.ms-powerpoint
.ppa	|	application/vnd.ms-powerpoint
.pptx	|	application/vnd.openxmlformats-officedocument.presentationml.presentation
.potx	|	application/vnd.openxmlformats-officedocument.presentationml.template
.ppsx	|	application/vnd.openxmlformats-officedocument.presentationml.slideshow
.ppam	|	application/vnd.ms-powerpoint.addin.macroEnabled.12
.pptm	|	application/vnd.ms-powerpoint.presentation.macroEnabled.12
.potm	|	application/vnd.ms-powerpoint.presentation.macroEnabled.12
.ppsm	|	application/vnd.ms-powerpoint.slideshow.macroEnabled.12

_The above information comes from [Office 2007 File Format MIME Types for HTTP Content Streaming](http://blogs.msdn.com/b/vsofficedeveloper/archive/2008/05/08/office-2007-open-xml-mime-types.aspx), it also mentioned:_
>  For Windows 2003 Servers running IIS 6.0, you can add the Open XML types in IIS Manager, Server Properties, MIME Types.  These new formats are included in Windows 2008 running IIS 7.0 by default.

## Acknowledgment

This [blog post](http://polymathprogrammer.com/2012/03/27/open-xml-sdk-class-structure/) helps me a lot, especially the diagram in it:
![OpenXML Class Structure](http://polymathprogrammer.com/images/blog/201203/openxmlsdkclassstructure.png)
