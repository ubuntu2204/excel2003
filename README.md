<!-- 
This README describes the package. If you publish this package to pub.dev,
this README's contents appear on the landing page for your package.

For information about how to write a good package README, see the guide for
[writing package pages](https://dart.dev/tools/pub/writing-package-pages). 

For general information about developing packages, see the Dart guide for
[creating packages](https://dart.dev/guides/libraries/create-packages)
and the Flutter guide for
[developing packages and plugins](https://flutter.dev/to/develop-packages). 
-->

A Dart library for reading legacy Excel (.xls) workbooks (Excel 97‑2003). It provides simple APIs to open a workbook, list sheets, and read cell values.

## Features

* Read BIFF8-formatted `.xls` files
* Iterate sheets, rows, and cells
* Access cell values/types and convert sheets to maps
* No external dependencies (pure Dart)

## Getting started

TODO: List prerequisites and provide or point to information on how to
start using the package.

## Usage

The following example demonstrates basic usage:

```dart
import 'package:excel2003/excel2003.dart';

void main() {
  final reader = XlsReader('path/to/workbook.xls');
  reader.open();
  print('Sheets: ${reader.sheetNames}');
  final sheet = reader.sheet(0);
  for (int row = sheet.firstRow; row < sheet.lastRow; row++) {
    for (int col = sheet.firstCol; col < sheet.lastCol; col++) {
      final value = sheet.cell(row, col);
      if (value != null) {
        print('Cell($row,$col)=$value');
      }
    }
  }
}

## Additional information

TODO: Tell users more about the package: where to find more information, how to 
contribute to the package, how to file issues, what response they can expect 
from the package authors, and more.
