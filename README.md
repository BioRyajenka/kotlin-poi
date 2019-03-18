# Kotlin-poi
Kotlin-poi is a convenient wrapper around apache poi. With this library you will find work with Excel is a peace of cake.

```kotlin
val wb = XSSFWorkbook()
wb.createSheetWithHeader(sheetName="Cakes", listOf("Cake name", "Yum level")) { sheet ->
	sheet.addRow("Butter Cake", 6)
	sheet.addRow("Sponge Cake", 7.5)
	sheet.addRow("Biscuit Cake", "10/10")
	
	sheet.addRow("Which one do you like most?")
	sheet.lastRow[1].createDropdownList("cakes_choices_hidden", listOf("Butter", "Sponge", "Biscuit"))
}
ExcelAdapter.write(wb, "lovely_cakes.xlsx")
```

## Things to do:
1. Add tests
2. Sort out poi-related hacks (:\)