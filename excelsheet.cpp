#include "excelsheet.h"

ExcelSheet::ExcelSheet(QAxObject* workSheet,int keyColumn)
{
    sheet = workSheet;
    key = keyColumn;
    usedRange = sheet->querySubObject("UsedRange");
    rows = usedRange->querySubObject("Rows");
    columns = usedRange->querySubObject("Columns");
    rowCount = rows->property("Count").toInt();
    colCount = columns->property("Count").toInt();
/**/
}

QVariantList ExcelSheet::getColume(int index){

  /**/  return usedRange->querySubObject("Columns(int)",index)
                    ->dynamicCall("value()")
                    .toList();


}

int ExcelSheet::writePerCol(int colSN,QVariantList colData){
      QAxObject *colume = usedRange->querySubObject("Columns(int)",colSN);
      if (colData.size()== rowCount){
         colume->setProperty("Value", colData);
     }
     return 0;
}
