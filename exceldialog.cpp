#include "exceldialog.h"
#include "excelsheet.h"
#include "ui_exceldialog.h"
#include"QDebug"
#include<QFileDialog>
#include<QMessageBox>

ExcelDialog::ExcelDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::ExcelDialog)
{
    ui->setupUi(this);


    QRegExp regExp("/S+");
    QRegExpValidator *expValidator = new QRegExpValidator(regExp,this);
    ui->Path_ObjectExcel->setValidator(expValidator);
    ui->Path_ReferenceExcel->setValidator(expValidator);
   // ui->Path_ObjectExcel->setText("D:/a-Qt5/forExcel2/ToNumber.xls");
  //  ui->Path_ReferenceExcel->setText("D:/a-Qt5/forExcel2/Number.xls");
  //  ui->spinBox_keyCol1->setValue(2);
  //  ui->spinBox_keyCol2->setValue(3);
  //  ui->spinBox_referCol->setValue(2);
 //   ui->spinBox_writeCol->setValue(1);
 //   ui->okButton->setEnabled(true);
    connect(ui->CancelButton,SIGNAL(clicked(bool)),this,SLOT(reject()));
    connect(ui->Path_ObjectExcel,SIGNAL(textChanged(QString)),this,SLOT(textChanged()));
    connect(ui->Path_ReferenceExcel,SIGNAL(textChanged(QString)),this,SLOT(textChanged()));
    connect(ui->okButton,SIGNAL(clicked(bool)),this,SLOT(saveValue()));
    connect(ui->pathButton,SIGNAL(clicked(bool)),this,SLOT(getObjFilePath()));
    connect(ui->pathButton2,SIGNAL(clicked(bool)),this,SLOT(getReferFilePath()));
    connect(ui->helpButton,SIGNAL(clicked(bool)),this,SLOT(help()));
}

ExcelDialog::~ExcelDialog()
{
    delete ui;
}


void ExcelDialog::textChanged()
{
   // bool b1=ui->Path_ObjectExcel->hasAcceptableInput();
  //  bool b2 = ui->Path_ReferenceExcel->hasAcceptableInput();
  //   qDebug()<<"value:"<<b1<<"&"<<b2;
  //  ui->okButton->setEnabled(b1&b2);
}

void ExcelDialog::saveValue()
{
    referenceFile = ui->Path_ReferenceExcel->text();
    objFile = ui->Path_ObjectExcel->text();
    keyColSN = ui->spinBox_keyCol2->value(); // 外键列
    referenceColSN =ui->spinBox_referCol->value() ; // 参考值列
    queryColSN = ui->spinBox_keyCol1->value(); //查询值列
    objColSN = ui->spinBox_writeCol->value(); //填充目标列
    handle();
    accept();
   // qDebug()<<referenceFile<<"obj:"<<objFile;
}

void ExcelDialog::handle(){
    QAxObject *excel = new QAxObject("Excel.Application");
    QAxObject *workbooks =  excel->querySubObject("WorkBooks");
    workbooks->dynamicCall( "Open (const QString&)", referenceFile);
    QAxObject *referenceBook = excel->querySubObject("ActiveWorkBook");
    QAxObject *rawReferenceSheet = referenceBook->querySubObject("Worksheets(int)", 1);

    workbooks->dynamicCall( "Open (const QString&)", objFile);
    QAxObject *objBook = excel->querySubObject("ActiveWorkBook");
    QAxObject *rawObjSheet = objBook->querySubObject("Worksheets(int)", 1);
    ExcelSheet *referenceSheet= new ExcelSheet(rawReferenceSheet,keyColSN);
    ExcelSheet *objSheet= new ExcelSheet(rawObjSheet,0);

    QVariantList nothing;
    nothing.push_back("Null");
    QVariantList keyCol = referenceSheet->getColume(keyColSN);
    QVariantList referenceCol = referenceSheet->getColume(referenceColSN);
    QVariantList queryCol = objSheet->getColume(queryColSN);
    int count = queryCol.size();

    QVariantList objColData;
    for(int i=0;i<count;i++){
        int index = keyCol.indexOf(queryCol[i]);

        if (index>=0){
             objColData.push_back(referenceCol[index]);

        }

        else{

            objColData.push_back(nothing);
        }

    }
    objSheet->writePerCol(objColSN,objColData);
    objBook->dynamicCall("Save()");
    objBook->dynamicCall("Close(Boolean)", false);
    referenceBook->dynamicCall("Close(Boolean)", false);
    excel->dynamicCall("Quit(void)");  //退出


}


void ExcelDialog::getObjFilePath(){
    QString filter = "Excel File (*.xls *.xlsx)";

    QString file_path =  QFileDialog::getOpenFileName(this,"选择目标文件...","./",filter);
        if(file_path.isEmpty())
        {
            return;
        }else{
           ui->Path_ObjectExcel->setText(file_path);
        }
}

void ExcelDialog::getReferFilePath(){ 
     QString filter = "Excel File (*.xls *.xlsx)";
     QString file_path =  QFileDialog::getOpenFileName(this,"选择参考文件...","./",filter);
        if(file_path.isEmpty())
        {
            return;
        }else{
           ui->Path_ReferenceExcel->setText(file_path);
        }

}

void ExcelDialog::help(){
     QMessageBox::information(this,tr("Help"),tr("<pre>功能:从参考表查询数据填入目标表（只操作excel第一个表).</pre>"
                                                          "<pre>操作: 1 需选择好Object File和Reference File .</pre>"
                                                          "<pre>     2 指定目标表的Key Colume和write Colume .</pre>"
                                                          "<pre>     3 参考表的Key Colume和Reference Colume .</pre>")
                                     ,QMessageBox::Close);
    return;
}
