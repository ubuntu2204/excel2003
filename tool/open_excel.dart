import 'package:excel2003/excel2003.dart';

void main() {
  print('正在打开 excel.xls...\n');

  final reader = XlsReader('excel.xls');
  reader.open();

  print('✓ 文件打开成功！');
  print('工作表数量: ${reader.sheetCount}');
  print('工作表名称: ${reader.sheetNames}');
  print('');

  // 遍历所有工作表
  for (int i = 0; i < reader.sheetCount; i++) {
    final sheet = reader.sheet(i);
    print('=== 工作表 ${i + 1}: "${sheet.name}" ===');
    print('行数: ${sheet.rowCount}');
    print('列数: ${sheet.colCount}');
    print(
      '数据范围: 行 ${sheet.firstRow}-${sheet.lastRow}, 列 ${sheet.firstCol}-${sheet.lastCol}',
    );
    print('');

    // 显示前10行数据
    final displayRows = sheet.rowCount < 10 ? sheet.rowCount : 10;
    print('前 $displayRows 行数据:');

    for (int row = sheet.firstRow; row < sheet.firstRow + displayRows; row++) {
      final values = <String>[];
      for (int col = sheet.firstCol; col < sheet.lastCol; col++) {
        final value = sheet.cell(row, col);
        values.add(value?.toString() ?? '');
      }
      print('  行 $row: ${values.join(' | ')}');
    }

    // 统计单元格类型
    final typeCount = <String, int>{};
    for (final cell in sheet.cells) {
      final typeName = cell.type.toString().split('.').last;
      typeCount[typeName] = (typeCount[typeName] ?? 0) + 1;
    }
    print('');
    print('单元格类型统计: $typeCount');
    print('');
    print('-' * 50);
    print('');
  }

  // 如果有数据，尝试转换为Map
  if (reader.sheetCount > 0) {
    final sheet = reader.sheet(0);
    if (sheet.rowCount > 1) {
      print('\n=== 第一行作为表头的数据 ===');
      final data = sheet.toMaps();
      final displayCount = data.length < 5 ? data.length : 5;
      print('显示前 $displayCount 条记录:');
      for (int i = 0; i < displayCount; i++) {
        print('  记录 ${i + 1}: ${data[i]}');
      }
      if (data.length > displayCount) {
        print('  ... 还有 ${data.length - displayCount} 条记录');
      }
    }
  }

  print('\n✓ 文件读取完成！');
}
