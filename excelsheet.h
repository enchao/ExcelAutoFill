#ifndef EXCELSHEET_H
#define EXCELSHEET_H

#include <ActiveQt/QAxWidget>
#include <ActiveQt/QAxObject>
#include "windows.h"

class ExcelSheet
{
public:
    ExcelSheet(QAxObject* workSheet,int keyColumn );
    QVariantList getColume(int index);
    int writePerCol(int colSN,QVariantList colData);
    ~ExcelSheet();
//private:
    QAxObject * sheet;
    QAxObject * rows;
    QAxObject * columns;   
    QAxObject * usedRange;
    int key;
    int rowCount;
    int colCount;

};

#endif // EXCELSHEET_H
