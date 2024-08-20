#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "clsexcelopt.h"
#include "clsxmlopt.h"
#include <QDebug>
#include <QFileInfo>
#include <QFileDialog>
#include <QProgressBar>
#include <QGridLayout>
#include "formprogressbar.h"
#include <QThread>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    this->setWindowTitle("Excel填充Ts");
    ui->widget_excel->setLayout(new QGridLayout(this));
    ui->widget_excel->layout()->setMargin(0);
    ui->widget_excel->hide();
    ui->widget_excel->setMaximumHeight(50);
    ui->widget_ts->setLayout(new QGridLayout(this));

}

MainWindow::~MainWindow()
{
    delete ui;
}

FormProgressBar* MainWindow::initProgressBar(FormProgressBar* _Bar, const QString& _name, bool _blerr)
{
    if(!_Bar)
    {
        _Bar = new FormProgressBar(this);
    }
    _Bar->setName(_name);
    if(_blerr)
    {
        _Bar->setName(_name + " excel not find!!!!!");
        _Bar->setError();
        _Bar->setProgress(100);
    }
    return _Bar;
}

clsExcelOpt* MainWindow::getExcelOpt()
{
    if(!m_pexcel)
    {
        m_pexcel = new clsExcelOpt();
        connect(m_pexcel, &clsExcelOpt::signalReadProcess, this, &MainWindow::slot_readExcelProcess);
    }
    return m_pexcel;
}

void MainWindow::readExcel()
{
    QString fileName = QFileDialog::getOpenFileName(this, tr("Open xls"), "C:/Users/E/Desktop", tr("Image Files (*.xls *.xlsx)"));
    if(fileName.isEmpty())
    {
        return;
    }

    m_pProgressExcel = initProgressBar(m_pProgressExcel, fileName);
    ui->widget_excel->layout()->addWidget(m_pProgressExcel);
    ui->widget_excel->show();
    getExcelOpt()->readCurSheet(fileName);
}

clsXmlOpt* MainWindow::getXmlOpt()
{
    if(!m_xmlTx)
    {
        m_xmlTx = new clsXmlOpt(this);
        connect(m_xmlTx, &clsXmlOpt::signal_loadTsProcess, this, &MainWindow::slot_loadTsProcess);
        connect(m_xmlTx, &clsXmlOpt::signal_importTsProcess, this, &MainWindow::slot_importTsProcess);
    }
    return m_xmlTx;
}

void MainWindow::starfillTs()
{
    for(auto task : m_mapTask)
    {//析构进度条
        delete task;
    }

    m_mapTask.clear();

    for(auto path : m_lststrPath)
    {
        QFileInfo ts(path);
        QString lang = ts.baseName().right(2);
        QMap<wordId, trans> mapExcleData = getExcelOpt()->getExcelData(lang);

        FormProgressBar* pBar = initProgressBar(new FormProgressBar(this), "import" + path, mapExcleData.isEmpty());
        int key = 0;
        if(mapExcleData.isEmpty())
        {
            key = getXmlOpt()->newKey();
        }
        else
        {
            key = getXmlOpt()->addWordData(path, mapExcleData);
        }
        ui->widget_ts->layout()->addWidget(pBar);
        m_mapTask.insert(key, pBar);
    }

}

void MainWindow::slot_readExcelProcess(int _process, bool _finish)
{
    _finish;
    if(m_pProgressExcel)
    {
        m_pProgressExcel->setProgress(_process);
    }
}

void MainWindow::slot_importTsProcess(int key, int _process, bool _finish)
{
    auto itor = m_mapTask.find(key);
    if(itor == m_mapTask.end())
    {
        return;
    }
    FormProgressBar* pBar = itor.value();
    pBar->setProgress(_process);

    if(_finish)
    {

    }
}

void MainWindow::slot_loadTsProcess(int key, int _process, bool _finish)
{
    auto itor = m_mapTask.find(key);
    if(itor == m_mapTask.end())
    {
        return;
    }
    FormProgressBar* pBar = itor.value();
    pBar->setProgress(_process);

    if(_finish){

    }

}

void MainWindow::on_pushButton_loadexcel_clicked()
{
    readExcel();
}

void MainWindow::on_pushButton_selectTsfile_clicked()
{
    m_lststrPath.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);
    dialog.setNameFilter(tr("Images (*.ts)"));
    dialog.setViewMode(QFileDialog::Detail);
    dialog.setFileMode(QFileDialog::ExistingFiles);
    QStringList fileNames;
    if (dialog.exec())
        fileNames = dialog.selectedFiles();
    qDebug()<<"fileNames:"<<fileNames;
    if(fileNames.isEmpty())
    {
        return;
    }

    for(auto tsName : fileNames)
    {
        m_lststrPath.push_back(tsName);
    }

    //读取ts
    loadTs();

}

void MainWindow::on_btn_excelts_clicked()
{
    getXmlOpt()->setReplace(Qt::Unchecked != ui->checkBox_excelts->checkState());
    starfillTs();
}

void MainWindow::on_btn_tsexcel_clicked()
{
    getExcelOpt()->setReplace(Qt::Unchecked != ui->checkBox_tsexcel->checkState());
    starfillExcel();
}

void MainWindow::starfillExcel()
{
    auto mapWordData = getXmlOpt()->getWordData();
    if(mapWordData.isEmpty()){
        qDebug()<<"MainWindow::starfillExcel mapWordData.isEmpty";
        return;
    }
    auto mapEmptyWordData = getXmlOpt()->getEmptyWordData();

    auto setLangs = getXmlOpt()->getLangs();
    getExcelOpt()->addTsData(mapWordData, mapEmptyWordData, setLangs);
}

void MainWindow::loadTs()
{
    for(auto task : m_mapTask)
    {//析构进度条
        delete task;
    }

    m_mapTask.clear();

    //读取ts
    for(auto path : m_lststrPath)
    {
        FormProgressBar* pBar = initProgressBar(new FormProgressBar(this), "load" + path);
        int key = 0;
        key = getXmlOpt()->addLoadTsTask(path);

        ui->widget_ts->layout()->addWidget(pBar);
        m_mapTask.insert(key, pBar);
    }
}

void MainWindow::on_pushButton_reload_clicked()
{
    QString fileName = getExcelOpt()->getCurPath();
    m_pProgressExcel = initProgressBar(m_pProgressExcel, fileName);
    ui->widget_excel->layout()->addWidget(m_pProgressExcel);
    ui->widget_excel->show();
    getExcelOpt()->readCurSheet(fileName);
}
