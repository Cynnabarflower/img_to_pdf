import 'dart:ffi';
import 'dart:io';
import 'dart:math';
import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:pdf/pdf.dart';
import 'package:pdf/widgets.dart' as pw;
import 'package:image/image.dart' as imglib;

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Img to Excel',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: const MyHomePage(),
    );
  }
}

List<String> oprosFiles = ['opros.jpg'];
String filename = 'images.pdf';
String message = '';
int maxImageWidth = 720;
int maxFileSizeKb = 1000;

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key});

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  List<Params> params = [];
  bool loading = false;

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: <Widget>[
            Expanded(
              child: loading
                  ? Center(
                      child: Column(
                        mainAxisSize: MainAxisSize.max,
                        mainAxisAlignment: MainAxisAlignment.center,
                        children: [
                          CircularProgressIndicator(),
                          SizedBox(
                            height: 8.0,
                          ),
                          Text('Loading...'),
                          Text(message)
                        ],
                      ),
                    )
                  : ListView(
                      children: [
                        ...params.map(
                          (e) => ListTile(
                            title: Text(e.toString()),
                            tileColor: e.error == 'done'
                                ? Colors.green.withOpacity(0.3)
                                : e.error == null
                                    ? Colors.white
                                    : Colors.redAccent.withOpacity(0.3),
                            subtitle: Text(e.error ?? e.message ?? ''),
                          ),
                        )
                      ],
                    ),
            ),
            Padding(
              padding: const EdgeInsets.all(32.0),
              child: Row(
                mainAxisAlignment: MainAxisAlignment.start,
                children: [
                  Expanded(
                    flex: 8,
                    child: Row(
                      mainAxisAlignment: MainAxisAlignment.start,
                      children: [
                        MaterialButton(
                          onPressed: () async {
                            final result = await FilePicker.platform.pickFiles(
                              dialogTitle: 'Excel file',
                              withData: true,
                            );
                            if (result == null) {
                              return;
                            }
                            params = [];
                            var bytes = result.files.first.bytes!;
                            return readFile(bytes).onError((error, stackTrace) {
                              print(error);
                              print(stackTrace);
                              setState(() {
                                loading = false;
                              });
                            });
                          },
                          child: Text('Выбрать excel файл'),
                        ),
                        MaterialButton(
                          onPressed: process,
                          child: Text('Запуск'),
                        ),
                      ],
                    ),
                  ),
                  Container(
                    width: 200,
                    child: TextField(
                      decoration: InputDecoration(labelText: 'Макс ширина изображения'),
                      controller:
                          TextEditingController(text: maxImageWidth.toString()),
                      onChanged: (v) {
                        maxImageWidth = int.tryParse(v) ?? maxImageWidth;
                        if (maxImageWidth <= 0) {
                          maxImageWidth = 99999;
                        }
                      },
                    ),
                  ),
                  Container(
                    width: 200,
                    child: TextField(
                      decoration: InputDecoration(labelText: 'Макс размер файла КБ'),
                      controller:
                          TextEditingController(text: maxFileSizeKb.toString()),
                      onChanged: (v) {
                        maxFileSizeKb = int.tryParse(v) ?? maxImageWidth;
                        if (maxFileSizeKb <= 0) {
                          maxFileSizeKb = 1000;
                        }
                      },
                    ),
                  ),
                ],
              ),
            )
          ],
        ),
      ),
    );
  }

  Future<pw.Document> _createDocument(
      Iterable<File> images, double compression) async {
    final pdf = pw.Document();
    for (final image in images) {
      Future<pw.Page> _imagePage() async {
        var im = (await imglib.decodeJpg(image.readAsBytesSync()))!;
        if (im.width > maxImageWidth && im.width < im.height) {
          im = imglib.copyResize(im, width: maxImageWidth);
        } else if (im.height > maxImageWidth && im.height < im.width) {
          im = imglib.copyResize(im, height: maxImageWidth);
        }
        return pw.Page(
          build: (pw.Context context) {
            var bytes =
                imglib.encodeJpg(im, quality: (100 * compression).floor());
            return pw.Image(
              pw.MemoryImage(
                bytes,
              ),
              width: im.width.toDouble(),
              height: im.height.toDouble(),
            ); // Center
          },
          pageFormat: PdfPageFormat(
            im!.width.toDouble(),
            im.height.toDouble(),
          ),
        );
      }

      pdf.addPage(await _imagePage());
    }
    return pdf;
  }

  Future<void> process() async {
    for (final param in params) {
      param.error = null;
      try {
        final dir = Directory(param.dir);
        List<File> images = [];
        List<File> otherFiles = [];
        await dir.list(recursive: true).forEach((element) {
          if (element is File) {
            final imagePaths = ['jpg'];
            if (oprosFiles.any((ex) => element.path.endsWith(ex))) {
              if (!Directory(param.savePath2).existsSync()) {
                Directory(param.savePath2).createSync(recursive: true);
              }
              element.copySync(
                '${param.savePath2}/${element.path.replaceAll('\\', '/').split('/').last}',
              );
            } else if (imagePaths
                .any((path) => element.path.toLowerCase().endsWith(path))) {
              if (element.path.replaceAll('\\', '/').split('/').last == 'VD.jpg') {
                if (!Directory(param.savePath).existsSync()) {
                  Directory(param.savePath).createSync(recursive: true);
                }
                element.copySync(
                  '${param.savePath}/${element.path.replaceAll('\\', '/').split('/').last}',
                );
                return;
              }
              images.add(element);
            } else {
              otherFiles.add(element);
            }
          }
        });

        if (!Directory(param.savePath).existsSync()) {
          Directory(param.savePath).createSync(recursive: true);
        }

        int docIndex = 0;
        int totalImages = images.length;
        setState(() {
          param.message = 'Загрузка ${totalImages} jpg изображений...';
        });
        await Future.delayed(Duration(milliseconds: 50));
        while (images.isNotEmpty) {
          int n = min(5, images.length);
          pw.Document doc =
              await _createDocument(images.take(n), param.compression);
          var bytes = await doc.save();
          if (bytes.length < maxFileSizeKb*1000) {
            n = min(10, images.length);
            doc = await _createDocument(images.take(n), param.compression);
            bytes = await doc.save();
          }
          var err = false;
          while (bytes.length > maxFileSizeKb*1000) {
            n--;
            if (n < 0) {
              err = true;
              break;
            }
            doc = await _createDocument(images.take(n), param.compression);
            bytes = await doc.save();
          }
          if (err) {
            setState(() {
              param.message =
              'Ошибка попробуйте уменьшить качество изображений или максимальный размер';
            });
            docIndex++;
            break;
          }
          images = images.sublist(n);
          setState(() {
            param.message =
                'Осталось ${images.length}/$totalImages jpg изображений...';
          });
          await Future.delayed(Duration(milliseconds: 50));

          File('${param.savePath}/${param.filename}${docIndex > 0 ? '$docIndex' : ''}.pdf')
              .writeAsBytesSync(bytes);
          docIndex++;
        }

        otherFiles.forEach((other) {
          other.copySync(
            '${param.savePath}/${other.path.replaceAll('\\', '/').split('/').last}',
          );
        });

        final row = param.data;
        if (row.length > 5) {
          bool createNSi = row[5]?.value.toString() == '1';
          if (createNSi) {
            await processNsi(row, param.savePath);
          }
        }

        param.error = 'done';
      } catch (e, s) {
        param.error = e.toString();
        print(e);
        print(s);
      }
      setState(() {});
    }
    setState(() {});
  }

  Future<void> processNsi(List<Data?> row, String savePath) async {
    final nsiTemplate = Excel.decodeBytes(
        (await rootBundle.load('assets/NSI_template.xlsx'))
            .buffer
            .asUint8List());
    int rowIndex = 0;
    final sheet = nsiTemplate.tables.values.first;
    for (final nsiRow in nsiTemplate.tables.values.first.rows) {
      if (nsiRow.isNotEmpty && nsiRow.first?.value.toString() == '\$') {
        for (int x = 0; x < row.length - 6; x++) {
          var val = row[x + 6]?.value;
          if (val is SharedString) {
            val = SharedString(node: val.node.copy());
          } else {
            print(val.runtimeType);
          }
          // if (nsiRow[x] != null) {
          //   nsiRow[x]!.value = val;
          // } else {
          //   nsiRow.add(Data.newData(sheet, rowIndex, x)..value = val);
          // }
          sheet.updateCell(
            CellIndex.indexByColumnRow(columnIndex: x, rowIndex: rowIndex),
            val,
            cellStyle: row[x + 6]?.cellStyle,
          );
          // nsiRow[x] = row[x + 6]?.value;
        }
        break;
      }
      rowIndex++;
    }

    var fileBytes = nsiTemplate.save()!;
    File('$savePath/NSI.xlsx')
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }

  Future<void> readFile(Uint8List bytes) async {
    setState(() {
      loading = true;
    });
    var excel = Excel.decodeBytes(bytes);
    for (var table in excel.tables.keys) {
      print(table); //sheet Name
      print(excel.tables[table]!.maxColumns);
      print(excel.tables[table]!.maxRows);
      for (var row in excel.tables[table]!.rows) {
        final dir = row.first?.value.toString();
        final fileName = row[1]?.value.toString();
        final compressionString = row[2]?.value.toString();
        final savePath = row[3]?.value.toString();
        final savePath2 = row[4]?.value.toString();

        if (dir == null ||
            fileName == null ||
            savePath == null ||
            savePath2 == null) {
          continue;
        }
        double compression = 1.0;
        if (compressionString != null && compressionString.contains('%')) {
          final aStr = compressionString.replaceAll(new RegExp(r'[^0-9.]'), '');
          compression = (double.tryParse(aStr) ?? 100) / 100.0;
        } else if (compressionString != null) {
          final aStr = compressionString.replaceAll(new RegExp(r'[^0-9.]'), '');
          compression = double.tryParse(aStr) ?? 1.0;
        }

        params.add(
          Params(
              dir: dir,
              filename: fileName,
              compression: compression,
              savePath: savePath,
              savePath2: savePath2,
              data: row),
        );
        print('$dir $fileName $compression $savePath $savePath2');
      }
      setState(() {
        loading = false;
      });
    }
  }
}

class Params {
  String dir;
  String filename;
  double compression;
  String savePath;
  String savePath2;
  String? error;
  String? message;
  List<Data?> data;

  Params({
    required this.dir,
    required this.filename,
    required this.compression,
    required this.savePath,
    required this.savePath2,
    required this.data,
  }) {
    this.dir = dir.replaceAll('\\', '/');
    this.savePath = savePath.replaceAll('\\', '/');
    this.savePath2 = savePath2.replaceAll('\\', '/');
    if (filename.contains('.pdf')) {
      filename = filename.replaceAll('.pdf', '');
    }
  }

  @override
  String toString() {
    return '$dir $filename $compression $savePath $savePath2';
  }
}
