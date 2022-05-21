import 'dart:io';
import 'package:flutter/material.dart' ;
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as excel;
import 'package:path_provider/path_provider.dart';
import 'package:open_file/open_file.dart';

class Ecxel extends StatefulWidget {
  const Ecxel({Key? key}) : super(key: key);

  @override
  State<Ecxel> createState() => _EcxelState();
}

class _EcxelState extends State<Ecxel> {
  var firstnameController = TextEditingController();
  var lastnameController = TextEditingController();
  var ageController = TextEditingController();
  var saveController = TextEditingController();

  @override
  Widget build(BuildContext context) {

    return Scaffold(
      appBar: AppBar(),
      body: Column(
        children: [
          TextFormField(
            controller: firstnameController,
            decoration: const InputDecoration(
              label: Text('First Name')
            ),
          ),
          TextFormField(
            controller: lastnameController,
            decoration:const InputDecoration(
                label: Text('Last Name')
            ),
          ),
          TextFormField(
            controller: ageController,
            decoration: const InputDecoration(
                label: Text('Age')
            ),
          ),
          TextFormField(
            controller: saveController,
            decoration: const InputDecoration(
                label: Text('saving')
            ),
          ),


          Center(
            child: ElevatedButton(
              child: const Text('طباعة تقرير'),
              onPressed: createdExcel,
            ),
          ),
        ],
      ) ,
    ) ;
  }

  Future<void> createdExcel() async{
    final excel.Workbook workbook = excel.Workbook();
    final excel.Worksheet sheet = workbook.worksheets[0] ;
    sheet.getRangeByName('A1').setText('First Name');
    sheet.getRangeByName('B1').setText('Last Name ');
    sheet.getRangeByName('C1').setText('Age');
    sheet.getRangeByName('D1').setText('save');

    sheet.getRangeByName('A2').setText(firstnameController.text);
    sheet.getRangeByName('B2').setText(lastnameController.text);
    sheet.getRangeByName('C2').setText(ageController.text);
    sheet.getRangeByName('D2').setText(saveController.text);

    final List<int> bytes = workbook.saveAsStream() ;
    workbook.dispose();


    final String path = (await getApplicationDocumentsDirectory()).path;
    final String fileName = '$path/Output.xlsx';
    final File file = File(fileName);
    await file.writeAsBytes(bytes , flush: true);
    OpenFile.open(fileName);

  }
}

