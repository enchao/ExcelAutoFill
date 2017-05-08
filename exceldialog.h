#ifndef EXCELDIALOG_H
#define EXCELDIALOG_H

#include <QDialog>

namespace Ui {
class ExcelDialog;
}

class ExcelDialog : public QDialog
{
    Q_OBJECT

public:
    explicit ExcelDialog(QWidget *parent = 0);
    ~ExcelDialog();
    void handle();


    QString referenceFile;
    QString objFile;
    int keyColSN; // 外键列
    int referenceColSN; // 参考值列

    int queryColSN ; //查询值列
    int objColSN; //填充目标列

private:
    Ui::ExcelDialog *ui;
private slots:
    void textChanged();
    void saveValue();
    void getObjFilePath();
    void getReferFilePath();
    void help();
};

#endif // EXCELDIALOG_H
