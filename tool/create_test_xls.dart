import 'dart:io';
import 'dart:typed_data';

/// Creates a minimal valid XLS file for testing
void main() {
  // This creates a minimal BIFF8 workbook with one sheet and a few cells
  final builder = XlsBuilder();
  builder.addSheet('Sheet1', [
    ['Name', 'Age', 'City'],
    ['Alice', 25, 'Beijing'],
    ['Bob', 30, 'Shanghai'],
    ['Charlie', 35, 'Guangzhou'],
  ]);

  final bytes = builder.build();
  File('test_data.xls').writeAsBytesSync(bytes);
  print('Created test_data.xls (${bytes.length} bytes)');
}

/// Simple XLS file builder for creating test files
class XlsBuilder {
  final List<_SheetData> _sheets = [];

  void addSheet(String name, List<List<dynamic>> data) {
    _sheets.add(_SheetData(name, data));
  }

  Uint8List build() {
    final output = BytesBuilder();

    // Build workbook stream first
    final workbook = _buildWorkbook();

    // OLE2 Header (512 bytes)
    final header = _buildOle2Header();
    output.add(header);

    // Calculate sectors needed for workbook
    final sectorSize = 512;
    final workbookSectors = (workbook.length + sectorSize - 1) ~/ sectorSize;

    // Pad workbook to sector boundary
    final paddedWorkbook = Uint8List(workbookSectors * sectorSize);
    paddedWorkbook.setRange(0, workbook.length, workbook);
    output.add(paddedWorkbook);

    // Add FAT sector
    final fat = _buildFat(workbookSectors);
    output.add(fat);

    // Add Directory sector
    final directory = _buildDirectory(workbook.length);
    output.add(directory);

    // Update header with correct values
    final result = Uint8List.fromList(output.toBytes());
    _updateHeader(result, workbookSectors);

    return result;
  }

  Uint8List _buildOle2Header() {
    final header = Uint8List(512);

    // Magic number
    header.setAll(0, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);

    // Minor version
    _writeUint16(header, 24, 0x003E);
    // Major version (3 = 512-byte sectors)
    _writeUint16(header, 26, 0x0003);
    // Byte order (little-endian)
    _writeUint16(header, 28, 0xFFFE);
    // Sector size power (9 = 512)
    _writeUint16(header, 30, 0x0009);
    // Mini sector size power (6 = 64)
    _writeUint16(header, 32, 0x0006);
    // Total sectors in FAT
    _writeUint32(header, 44, 1);
    // First directory sector (will be updated)
    _writeUint32(header, 48, 0);
    // Mini stream cutoff size
    _writeUint32(header, 56, 0x1000);
    // First mini FAT sector (none)
    _writeUint32(header, 60, 0xFFFFFFFE);
    // Mini FAT sector count
    _writeUint32(header, 64, 0);
    // First DIFAT sector (none)
    _writeUint32(header, 68, 0xFFFFFFFE);
    // DIFAT sector count
    _writeUint32(header, 72, 0);

    // First FAT sector in DIFAT (will be updated)
    _writeUint32(header, 76, 0);

    // Fill rest of DIFAT with 0xFFFFFFFE
    for (int i = 80; i < 512; i += 4) {
      _writeUint32(header, i, 0xFFFFFFFE);
    }

    return header;
  }

  Uint8List _buildWorkbook() {
    final output = BytesBuilder();

    // BOF - Workbook globals
    output.add(_buildBof(0x0005));

    // CODEPAGE
    output.add(_buildCodepage());

    // Collect all strings for SST
    final strings = <String>[];
    final stringIndex = <String, int>{};
    for (final sheet in _sheets) {
      for (final row in sheet.data) {
        for (final cell in row) {
          if (cell is String && !stringIndex.containsKey(cell)) {
            stringIndex[cell] = strings.length;
            strings.add(cell);
          }
        }
      }
    }

    // BOUNDSHEET records (one per sheet)
    // We need to calculate the offset to each sheet's BOF
    // For simplicity, we'll update these after building the globals
    final boundsheetPositions = <int>[];
    for (final sheet in _sheets) {
      boundsheetPositions.add(output.length);
      output.add(_buildBoundsheet(sheet.name, 0)); // Placeholder offset
    }

    // SST - Shared String Table
    if (strings.isNotEmpty) {
      output.add(_buildSst(strings));
    }

    // EOF for globals
    output.add(_buildEof());

    // Now build each sheet and update BOUNDSHEET offsets
    final sheetOffsets = <int>[];
    for (int i = 0; i < _sheets.length; i++) {
      sheetOffsets.add(output.length);
      output.add(_buildSheet(_sheets[i], stringIndex));
    }

    // Build final bytes and update BOUNDSHEET records
    final bytes = Uint8List.fromList(output.toBytes());
    for (int i = 0; i < _sheets.length; i++) {
      final pos = boundsheetPositions[i] + 4; // Skip record header
      _writeUint32(bytes, pos, sheetOffsets[i]);
    }

    return bytes;
  }

  Uint8List _buildBof(int type) {
    final record = Uint8List(20);
    _writeUint16(record, 0, 0x0809); // BOF
    _writeUint16(record, 2, 16); // Length
    _writeUint16(record, 4, 0x0600); // BIFF8
    _writeUint16(record, 6, type); // Type
    _writeUint16(record, 8, 0x0DBB); // Build ID
    _writeUint16(record, 10, 0x07CC); // Build year
    _writeUint32(record, 12, 0); // File history
    _writeUint32(record, 16, 0x00000006); // Lowest BIFF version
    return record;
  }

  Uint8List _buildCodepage() {
    final record = Uint8List(6);
    _writeUint16(record, 0, 0x0042); // CODEPAGE
    _writeUint16(record, 2, 2); // Length
    _writeUint16(record, 4, 0x04E4); // UTF-16
    return record;
  }

  Uint8List _buildBoundsheet(String name, int offset) {
    final nameBytes = _encodeString(name);
    final record = Uint8List(12 + nameBytes.length);
    _writeUint16(record, 0, 0x0085); // BOUNDSHEET
    _writeUint16(record, 2, 8 + nameBytes.length); // Length
    _writeUint32(record, 4, offset); // BOF offset
    record[8] = 0; // Visibility (visible)
    record[9] = 0; // Sheet type (worksheet)
    record[10] = name.length; // Name length
    record[11] = 0; // Option flags (8-bit chars)
    record.setRange(12, 12 + nameBytes.length, nameBytes);
    return record;
  }

  Uint8List _buildSst(List<String> strings) {
    final output = BytesBuilder();

    // Calculate total size
    int totalChars = 0;
    for (final s in strings) {
      totalChars += 3 + s.length; // length (2) + flags (1) + chars
    }

    // Header: total strings, unique strings
    final header = Uint8List(8);
    _writeUint32(header, 0, strings.length);
    _writeUint32(header, 4, strings.length);
    output.add(header);

    // Each string
    for (final s in strings) {
      final strRecord = Uint8List(3 + s.length);
      _writeUint16(strRecord, 0, s.length);
      strRecord[2] = 0; // 8-bit characters
      for (int i = 0; i < s.length; i++) {
        strRecord[3 + i] = s.codeUnitAt(i) & 0xFF;
      }
      output.add(strRecord);
    }

    final data = output.toBytes();
    final record = Uint8List(4 + data.length);
    _writeUint16(record, 0, 0x00FC); // SST
    _writeUint16(record, 2, data.length);
    record.setRange(4, 4 + data.length, data);
    return record;
  }

  Uint8List _buildEof() {
    final record = Uint8List(4);
    _writeUint16(record, 0, 0x000A); // EOF
    _writeUint16(record, 2, 0); // Length
    return record;
  }

  Uint8List _buildSheet(_SheetData sheet, Map<String, int> stringIndex) {
    final output = BytesBuilder();

    // BOF - Sheet
    output.add(_buildBof(0x0010));

    // DIMENSION
    output.add(
      _buildDimension(
        sheet.data.length,
        sheet.data.isEmpty ? 0 : sheet.data[0].length,
      ),
    );

    // Cell data
    for (int row = 0; row < sheet.data.length; row++) {
      for (int col = 0; col < sheet.data[row].length; col++) {
        final value = sheet.data[row][col];
        if (value is String) {
          final idx = stringIndex[value]!;
          output.add(_buildLabelSst(row, col, idx));
        } else if (value is num) {
          output.add(_buildNumber(row, col, value.toDouble()));
        }
      }
    }

    // EOF
    output.add(_buildEof());

    return Uint8List.fromList(output.toBytes());
  }

  Uint8List _buildDimension(int rows, int cols) {
    final record = Uint8List(18);
    _writeUint16(record, 0, 0x0200); // DIMENSION
    _writeUint16(record, 2, 14); // Length
    _writeUint32(record, 4, 0); // First row
    _writeUint32(record, 8, rows); // Last row + 1
    _writeUint16(record, 12, 0); // First col
    _writeUint16(record, 14, cols); // Last col + 1
    _writeUint16(record, 16, 0); // Reserved
    return record;
  }

  Uint8List _buildLabelSst(int row, int col, int sstIndex) {
    final record = Uint8List(14);
    _writeUint16(record, 0, 0x00FD); // LABELSST
    _writeUint16(record, 2, 10); // Length
    _writeUint16(record, 4, row); // Row
    _writeUint16(record, 6, col); // Column
    _writeUint16(record, 8, 0); // XF index
    _writeUint32(record, 10, sstIndex); // SST index
    return record;
  }

  Uint8List _buildNumber(int row, int col, double value) {
    final record = Uint8List(18);
    _writeUint16(record, 0, 0x0203); // NUMBER
    _writeUint16(record, 2, 14); // Length
    _writeUint16(record, 4, row); // Row
    _writeUint16(record, 6, col); // Column
    _writeUint16(record, 8, 0); // XF index

    // Write IEEE 754 double
    final bytes = ByteData(8);
    bytes.setFloat64(0, value, Endian.little);
    for (int i = 0; i < 8; i++) {
      record[10 + i] = bytes.getUint8(i);
    }
    return record;
  }

  Uint8List _buildFat(int workbookSectors) {
    final fat = Uint8List(512);

    // Workbook sectors chain
    for (int i = 0; i < workbookSectors - 1; i++) {
      _writeUint32(fat, i * 4, i + 1);
    }
    // End of workbook chain
    _writeUint32(fat, (workbookSectors - 1) * 4, 0xFFFFFFFE);

    // FAT sector (this sector)
    _writeUint32(fat, workbookSectors * 4, 0xFFFFFFFD);

    // Directory sector
    _writeUint32(fat, (workbookSectors + 1) * 4, 0xFFFFFFFE);

    // Fill rest with free
    for (int i = workbookSectors + 2; i < 128; i++) {
      _writeUint32(fat, i * 4, 0xFFFFFFFF);
    }

    return fat;
  }

  Uint8List _buildDirectory(int workbookSize) {
    final dir = Uint8List(512);

    // Root Entry (128 bytes)
    _writeDirectoryEntry(dir, 0, 'Root Entry', 5, 0xFFFFFFFE, 0);

    // Workbook entry (128 bytes)
    _writeDirectoryEntry(dir, 128, 'Workbook', 2, 0, workbookSize);

    // Fill rest with empty entries
    for (int i = 256; i < 512; i += 128) {
      _writeDirectoryEntry(dir, i, '', 0, 0, 0);
    }

    return dir;
  }

  void _writeDirectoryEntry(
    Uint8List data,
    int offset,
    String name,
    int type,
    int startSector,
    int size,
  ) {
    // Name (UTF-16LE)
    for (int i = 0; i < name.length && i < 31; i++) {
      _writeUint16(data, offset + i * 2, name.codeUnitAt(i));
    }
    // Name size (including null terminator)
    _writeUint16(data, offset + 64, (name.length + 1) * 2);
    // Type
    data[offset + 66] = type;
    // Color (black)
    data[offset + 67] = 1;
    // Left/Right/Child sibling IDs
    _writeUint32(data, offset + 68, 0xFFFFFFFF);
    _writeUint32(data, offset + 72, 0xFFFFFFFF);
    _writeUint32(data, offset + 76, type == 5 ? 1 : 0xFFFFFFFF);
    // Start sector
    _writeUint32(data, offset + 116, startSector);
    // Size
    _writeUint32(data, offset + 120, size);
  }

  void _updateHeader(Uint8List data, int workbookSectors) {
    // FAT sector location
    _writeUint32(data, 76, workbookSectors);
    // Directory sector location
    _writeUint32(data, 48, workbookSectors + 1);
  }

  Uint8List _encodeString(String s) {
    return Uint8List.fromList(s.codeUnits.map((c) => c & 0xFF).toList());
  }

  void _writeUint16(Uint8List data, int offset, int value) {
    data[offset] = value & 0xFF;
    data[offset + 1] = (value >> 8) & 0xFF;
  }

  void _writeUint32(Uint8List data, int offset, int value) {
    data[offset] = value & 0xFF;
    data[offset + 1] = (value >> 8) & 0xFF;
    data[offset + 2] = (value >> 16) & 0xFF;
    data[offset + 3] = (value >> 24) & 0xFF;
  }
}

class _SheetData {
  final String name;
  final List<List<dynamic>> data;

  _SheetData(this.name, this.data);
}
