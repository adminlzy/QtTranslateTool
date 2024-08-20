#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <list>
#include <QMap>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class clsExcelOpt;
class clsXmlOpt;
class FormProgressBar;
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();


private slots:

    void slot_readExcelProcess(int _process, bool _finish);
    void slot_importTsProcess(int key, int _process, bool _finish);
    void slot_loadTsProcess(int key, int _process, bool _finish);
    void on_pushButton_loadexcel_clicked();

    void on_pushButton_selectTsfile_clicked();

    void on_btn_excelts_clicked();

    void on_btn_tsexcel_clicked();

    void on_pushButton_reload_clicked();

private:
    void readExcel();
    void starfillTs();
    void starfillExcel();
    void loadTs();

    FormProgressBar* initProgressBar(FormProgressBar* _Bar, const QString& _name, bool _blerr=false);

    clsXmlOpt* getXmlOpt();
    clsExcelOpt* getExcelOpt();
private:
    Ui::MainWindow *ui;
    clsExcelOpt* m_pexcel = nullptr;
    clsXmlOpt* m_xmlTx = nullptr;
    std::list<QString> m_lststrPath;
    QMap<int, FormProgressBar*> m_mapTask;
    FormProgressBar* m_pProgressExcel = nullptr; //读取excel文件进度

};
#endif // MAINWINDOW_H
