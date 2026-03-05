import 'dart:io';

import 'package:excel2003/src/ole2/ole2_reader.dart';
import 'package:excel2003/src/biff/biff_records.dart';

void main() {
  final file = File('test_data.xls');
  if (!file.existsSync()) {
    print('File not found');
    return;
  }

  print('=== OLE2 Structure ===');
  final ole2 = Ole2Reader();
  ole2.open('test_data.xls');

  print('Directory entries:');
  for (final entry in ole2.entries) {
    print(
      '  ${entry.name}: type=${entry.type}, sector=${entry.startSector}, size=${entry.size}',
    );
  }

  final workbookEntry = ole2.findEntry('Workbook') ?? ole2.findEntry('Book');
  if (workbookEntry == null) {
    print('No Workbook stream found!');
    return;
  }

  print('\nWorkbook entry found: ${workbookEntry.name}');
  final workbookData = ole2.readEntry(workbookEntry);
  print('Workbook stream size: ${workbookData.length} bytes');

  print('\n=== BIFF Records ===');
  final parser = BiffParser(workbookData);
  int recordCount = 0;

  while (parser.hasMore && recordCount < 50) {
    final record = parser.nextRecord();
    if (record == null) break;

    String typeName;
    switch (record.type) {
      case BiffRecordType.bof:
        typeName = 'BOF';
        break;
      case BiffRecordType.eof:
        typeName = 'EOF';
        break;
      case BiffRecordType.boundSheet:
        typeName = 'BOUNDSHEET';
        break;
      case BiffRecordType.sst:
        typeName = 'SST';
        break;
      case BiffRecordType.dimension:
        typeName = 'DIMENSION';
        break;
      case BiffRecordType.number:
        typeName = 'NUMBER';
        break;
      case BiffRecordType.labelSst:
        typeName = 'LABELSST';
        break;
      case BiffRecordType.codePage:
        typeName = 'CODEPAGE';
        break;
      default:
        typeName = '0x${record.type.toRadixString(16).padLeft(4, '0')}';
    }

    print('  $typeName: ${record.length} bytes at offset ${record.offset}');

    if (record.type == BiffRecordType.bof) {
      final bof = BofRecord.fromRecord(record);
      print(
        '    -> version=0x${bof.version.toRadixString(16)}, type=0x${bof.type.toRadixString(16)}',
      );
    } else if (record.type == BiffRecordType.boundSheet) {
      final bs = BoundSheetRecord.fromRecord(record);
      print('    -> name="${bs.name}", bofPos=${bs.bofPosition}');
    }

    recordCount++;
  }

  print('\nTotal records scanned: $recordCount');
}
