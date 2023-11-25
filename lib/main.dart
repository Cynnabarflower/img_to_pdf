import 'dart:io';
import 'dart:math';
import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
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

List<String> excludeFiles = ['opros.jpg'];
String filename = 'images.pdf';
String message = '';

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
                      var bytes = result!.files.first.bytes!;
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
            )
          ],
        ),
      ),
    );
  }

  Future<pw.Document> _createDocument(
      Iterable<File> images, double imageScale, double compression) async {
    final pdf = pw.Document();
    for (final image in images) {
      Future<pw.Page> _imagePage() async {
        final im = await imglib.decodeJpg(image.readAsBytesSync());
        return pw.Page(
          build: (pw.Context context) {
            return pw.Image(
              pw.MemoryImage(
                imglib.encodeJpg(im!, quality: (100 * compression).floor()),
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
            if (imagePaths.any((path) => element.path.endsWith(path)) &&
                !excludeFiles.any((ex) => element.path.endsWith(ex))) {
              images.add(element);
            } else {
              otherFiles.add(element);
            }
          }
        });

        if (!Directory(param.savePath).existsSync()) {
          Directory(param.savePath).createSync(recursive: true);
        }

        if (param.savePath2 != null &&
            !Directory(param.savePath2!).existsSync()) {
          Directory(param.savePath2!).createSync(recursive: true);
        }

        int docIndex = 0;
        int totalImages = images.length;
        setState(() {
          param.message = 'Загрузка ${totalImages} jpg изображений...';
        });
        await Future.delayed(Duration(milliseconds: 50));
        while (images.isNotEmpty) {
          int n = min(5, images.length);
          pw.Document doc = await _createDocument(
              images.take(n), param.imageScale, param.compression);
          var bytes = await doc.save();
          if (bytes.length < 1000000) {
            n = min(10, images.length);
            doc = await _createDocument(
                images.take(n), param.imageScale, param.compression);
            bytes = await doc.save();
          }
          while (bytes.length > 1000000 && n > 0) {
            n--;
            doc = await _createDocument(
                images.take(n), param.imageScale, param.compression);
            bytes = await doc.save();
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
              '${param.savePath2}/${other.path.replaceAll('\\', '/').split('/').last}');
        });
        param.error = 'done';
      } catch (e) {
        param.error = e.toString();
      }
      setState(() {});
    }
    setState(() {});
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
        final imageScaleString =
            row.length >= 5 ? row[5]?.value.toString() : null;
        if (dir == null || fileName == null || savePath == null || savePath2 == null) {
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
        final imageScale = double.tryParse(imageScaleString ?? '');

        params.add(
          Params(
            dir: dir,
            filename: fileName,
            compression: compression,
            savePath: savePath,
            savePath2: savePath2,
            imageScale: imageScale ?? 1.0,
          ),
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
  double imageScale;

  Params({
    required this.dir,
    required this.filename,
    required this.compression,
    required this.savePath,
    required this.savePath2,
    this.imageScale = 1.0,
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
