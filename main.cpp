
#include <QApplication>
#include <QDebug>
#include "excelsheet.h"
#include "exceldialog.h"


//添加新文件时，记得执行qmake，否则会报这个错误

    int main(int argc, char *argv[])
    {
        QApplication a(argc, argv);      
        ExcelDialog w;
        w.show();
        return a.exec();
    }
