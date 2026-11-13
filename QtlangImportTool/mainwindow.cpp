#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "clsexcelopt.h"
#include "clsxmlopt.h"
#include <QDebug>
#include <QFileInfo>
#include <QFileDialog>
#include <QDir>
#include <QProgressBar>
#include <QGridLayout>
#include "formprogressbar.h"
#include <QThread>
#include <QPushButton>
#include <QSettings>
#include <QTimer>
#include <QUrl>
#include <QMimeData>
#include <QDragEnterEvent>
#include <QDropEvent>
#include "dropdownarea.h"
#include <QComboBox>
#include <QInputDialog>
#include <QLineEdit>
#include <QMessageBox>
#include <QStatusBar>

namespace {
const char* const kExcelFileKey = "recent/excelFile";
const char* const kTsFilesKey = "recent/tsFiles";
const char* const kTsDirKey = "recent/tsDir";
const char* const kProfileNameKey = "recent/profileName";
const char* const kProfilesGroup = "profiles";

bool isExcelFile(const QString& _path)
{
    const QString suffix = QFileInfo(_path).suffix().toLower();
    return suffix == "xls" || suffix == "xlsx";
}

bool isTsFile(const QString& _path)
{
    return QFileInfo(_path).suffix().toLower() == "ts";
}
}

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    this->setWindowTitle("Excel填充Ts");
    ui->widget_excel->setLayout(new QGridLayout(ui->widget_excel));
    ui->widget_excel->layout()->setContentsMargins(0, 0, 0, 0);
    ui->widget_excel->hide();
    ui->widget_excel->setMaximumHeight(50);
    ui->widget_ts->setLayout(new QGridLayout(ui->widget_ts));
    ui->widget_ts->layout()->setContentsMargins(0, 0, 0, 0);
    ui->widget_ts->layout()->setSpacing(8);

    //左侧文件区占主要宽度，右侧操作栏保持紧凑
    ui->mainLayout->setStretch(0, 3);
    ui->mainLayout->setStretch(1, 2);
    ui->layout_input->setStretch(5, 1);
    statusBar()->showMessage(tr("就绪"));

    //整个窗口也支持拖拽
    setAcceptDrops(true);
    connect(ui->dropArea, &DropArea::filesDropped, this, &MainWindow::slot_filesDropped);

    connect(ui->combo_profile, QOverload<int>::of(&QComboBox::currentIndexChanged), this, &MainWindow::slot_profileChanged);
    connect(ui->btn_saveProfile, &QPushButton::clicked, this, &MainWindow::slot_saveProfile);
    connect(ui->btn_deleteProfile, &QPushButton::clicked, this, &MainWindow::slot_deleteProfile);
    refreshProfileCombo();

    //窗口显示后恢复上次选择的文件
    QTimer::singleShot(0, this, &MainWindow::restoreLastSelection);
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
        _Bar->setError();
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
    QString startDir = QDir::homePath();
    if(!m_lastExcelPath.isEmpty())
    {
        startDir = QFileInfo(m_lastExcelPath).absolutePath();
    }

    QString fileName = QFileDialog::getOpenFileName(this, tr("Open xls"), startDir, tr("Excel Files (*.xls *.xlsx)"));
    if(fileName.isEmpty())
    {
        return;
    }

    loadExcelFile(fileName);
}

void MainWindow::loadExcelFile(const QString& _fileName)
{
    m_lastExcelPath = _fileName;

    if(!m_bLoadingProfile)
    {//手动换文件后当前选择已不再对应某个匹配项
        clearProfileSelection();
    }

    m_pProgressExcel = initProgressBar(m_pProgressExcel, _fileName);
    ui->widget_excel->layout()->addWidget(m_pProgressExcel);
    ui->widget_excel->show();

    saveSettings();
    getExcelOpt()->readCurSheet(_fileName);
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
#include <string>
#include <iostream>

struct leakMem{
    std::string m_str1 = "1231321321313132132132131313132131";
    std::string m_str2 = "1231321321313132132132131313132131";
    std::string m_str3 = "1231321321313132132132131313132131";
    std::string m_str4 = "1231321321313132132132131313132131";
    std::string m_str5 = "1231321321313132132132131313132131";
    std::string m_str6 = "1231321321313132132132131313132131";
    std::string m_str7 = "1231321321313132132132131313132131";
    std::string m_str11 = "1231321321313132132132131313132131";
    std::string m_str21 = "1231321321313132132132131313132131";
    std::string m_str31 = "1231321321313132132132131313132131";
    std::string m_str41 = "1231321321313132132132131313132131";
    std::string m_str51 = "1231321321313132132132131313132131";
    std::string m_str61 = "1231321321313132132132131313132131";
    std::string m_str71 = "1231321321313132132132131313132131";
    std::string m_str12 = "1231321321313132132132131313132131";
    std::string m_str22 = "1231321321313132132132131313132131";
    std::string m_str32 = "1231321321313132132132131313132131";
    std::string m_str42 = "1231321321313132132132131313132131";
    std::string m_str52 = "1231321321313132132132131313132131";
    std::string m_str62 = "1231321321313132132132131313132131";
    std::string m_str72 = "1231321321313132132132131313132131";
};

void MainWindow::createbtn(int _num){
//    QVector<leakMem*> vec;
//    for(int i = 0; i < _num; i++){
//        vec.push_back(new leakMem);
//    }
    const int STR_SIZE = 1024;
    char * pStr = new char [STR_SIZE];
    strcpy_s(pStr, STR_SIZE, "Memory Leak Sample");
    std::cout << pStr << std::endl;
    delete[] pStr;
}

void MainWindow::on_pushButton_selectTsfile_clicked()
{
//    createbtn(100);
    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);
    dialog.setNameFilter(tr("TS Files (*.ts)"));
    dialog.setViewMode(QFileDialog::Detail);
    dialog.setFileMode(QFileDialog::ExistingFiles);
    dialog.setDirectory(lastTsDir());
    QStringList fileNames;
    if (dialog.exec())
        fileNames = dialog.selectedFiles();
    qDebug()<<"fileNames:"<<fileNames;
    if(fileNames.isEmpty())
    {
        return;
    }

    setTsFiles(fileNames);
}

void MainWindow::setTsFiles(const QStringList& _fileNames)
{
    m_lststrPath.clear();
    for(const auto& tsName : _fileNames)
    {
        m_lststrPath.push_back(tsName);
    }

    if(!_fileNames.isEmpty())
    {
        m_lastTsDir = QFileInfo(_fileNames.first()).absolutePath();
    }

    if(!m_bLoadingProfile)
    {//手动换文件后当前选择已不再对应某个匹配项
        clearProfileSelection();
    }

    saveSettings();

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

    QString noTransPath = QFileDialog::getSaveFileName(this, tr("Save no translation file"), QDir::homePath(), tr("Excel Files (*.xlsx)"));
    if(noTransPath.isEmpty()){
        return;
    }
    getExcelOpt()->addTsData(mapWordData, mapEmptyWordData, setLangs, noTransPath);
}

void MainWindow::loadTs()
{
    for(auto task : m_mapTask)
    {//析构进度条
        delete task;
    }

    m_mapTask.clear();

    //新一批ts，先清空上一批的数据
    getXmlOpt()->clearTsData();

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
    if(fileName.isEmpty())
    {
        fileName = m_lastExcelPath;
    }
    if(fileName.isEmpty() || !QFileInfo::exists(fileName))
    {
        QMessageBox::information(this, tr("重新加载"), tr("当前没有可重新加载的excel文件。"));
        return;
    }
    m_pProgressExcel = initProgressBar(m_pProgressExcel, fileName);
    ui->widget_excel->layout()->addWidget(m_pProgressExcel);
    ui->widget_excel->show();
    getExcelOpt()->readCurSheet(fileName);
}

void MainWindow::saveSettings()
{
    QSettings settings(QSettings::IniFormat, QSettings::UserScope, "QtTranslateTool", "langImport");

    QStringList tsFiles;
    for(const auto& path : m_lststrPath)
    {
        tsFiles.push_back(path);
    }

    settings.setValue(kExcelFileKey, m_lastExcelPath);
    settings.setValue(kTsFilesKey, tsFiles);
    settings.setValue(kTsDirKey, m_lastTsDir);
}

void MainWindow::restoreLastSelection()
{
    loadProfiles();

    QSettings settings(QSettings::IniFormat, QSettings::UserScope, "QtTranslateTool", "langImport");

    m_lastExcelPath = settings.value(kExcelFileKey).toString();
    m_lastTsDir = settings.value(kTsDirKey).toString();
    m_curProfileName = settings.value(kProfileNameKey).toString();

    if(!m_curProfileName.isEmpty() && m_mapProfiles.contains(m_curProfileName))
    {//优先恢复上次使用的匹配项
        refreshProfileCombo(m_curProfileName);
        applyProfile(m_curProfileName, false);
        return;
    }

    m_curProfileName.clear();
    refreshProfileCombo();

    QStringList tsFiles;
    const QStringList savedTsFiles = settings.value(kTsFilesKey).toStringList();
    for(const QString& path : savedTsFiles)
    {
        if(QFileInfo::exists(path) && isTsFile(path))
        {
            tsFiles.push_back(path);
        }
    }

    if(!m_lastExcelPath.isEmpty() && isExcelFile(m_lastExcelPath) && QFileInfo::exists(m_lastExcelPath))
    {
        loadExcelFile(m_lastExcelPath);
    }
    else
    {
        m_lastExcelPath.clear();
    }

    if(!tsFiles.isEmpty())
    {
        setTsFiles(tsFiles);
    }

    saveSettings();
}

void MainWindow::loadProfiles()
{
    m_mapProfiles.clear();

    QSettings settings(QSettings::IniFormat, QSettings::UserScope, "QtTranslateTool", "langImport");
    settings.beginGroup(kProfilesGroup);
    const QStringList names = settings.value("names").toStringList();
    settings.endGroup();

    for(const QString& name : names)
    {
        if(name.isEmpty())
        {
            continue;
        }

        MatchProfile profile;
        settings.beginGroup(QString("%1/items/%2").arg(QString::fromLatin1(kProfilesGroup), name));
        profile.excelPath = settings.value("excelFile").toString();
        profile.tsFiles = settings.value("tsFiles").toStringList();
        profile.tsDir = settings.value("tsDir").toString();
        settings.endGroup();

        m_mapProfiles.insert(name, profile);
    }
}

void MainWindow::saveProfiles()
{
    QSettings settings(QSettings::IniFormat, QSettings::UserScope, "QtTranslateTool", "langImport");
    settings.remove(kProfilesGroup);
    if(m_mapProfiles.isEmpty())
    {//空列表会被写成@Invalid()，直接不落盘
        settings.sync();
        return;
    }

    settings.beginGroup(kProfilesGroup);
    settings.setValue("names", QStringList(m_mapProfiles.keys()));
    settings.endGroup();

    for(auto itor = m_mapProfiles.constBegin(); itor != m_mapProfiles.constEnd(); ++itor)
    {
        settings.beginGroup(QString("%1/items/%2").arg(QString::fromLatin1(kProfilesGroup), itor.key()));
        settings.setValue("excelFile", itor.value().excelPath);
        settings.setValue("tsFiles", itor.value().tsFiles);
        settings.setValue("tsDir", itor.value().tsDir);
        settings.endGroup();
    }

    settings.sync();
}

void MainWindow::refreshProfileCombo(const QString& _selected)
{
    const bool bLoading = m_bLoadingProfile;
    m_bLoadingProfile = true;

    ui->combo_profile->clear();
    ui->combo_profile->addItem(tr("（当前选择）"));
    for(auto itor = m_mapProfiles.constBegin(); itor != m_mapProfiles.constEnd(); ++itor)
    {
        ui->combo_profile->addItem(itor.key());
        ui->combo_profile->setItemData(ui->combo_profile->count() - 1,
                                       tr("Excel: %1\nTS(%2): %3")
                                       .arg(itor.value().excelPath)
                                       .arg(itor.value().tsFiles.size())
                                       .arg(itor.value().tsFiles.join("\n")),
                                       Qt::ToolTipRole);
    }

    int index = 0;
    if(!_selected.isEmpty())
    {
        index = ui->combo_profile->findText(_selected);
    }
    ui->combo_profile->setCurrentIndex(index < 0 ? 0 : index);

    m_bLoadingProfile = bLoading;
}

void MainWindow::applyProfile(const QString& _name, bool _interactive)
{
    auto itor = m_mapProfiles.find(_name);
    if(itor == m_mapProfiles.end())
    {
        return;
    }

    const MatchProfile profile = itor.value();
    const bool bLoading = m_bLoadingProfile;
    m_bLoadingProfile = true;
    m_curProfileName = _name;

    QStringList missing;
    QStringList tsFiles;
    for(const QString& path : profile.tsFiles)
    {
        if(QFileInfo::exists(path) && isTsFile(path))
        {
            tsFiles.push_back(path);
        }
        else
        {
            missing.push_back(path);
        }
    }

    if(!profile.excelPath.isEmpty() && isExcelFile(profile.excelPath) && QFileInfo::exists(profile.excelPath))
    {
        loadExcelFile(profile.excelPath);
    }
    else
    {
        showExcelError(profile.excelPath);
        missing.push_back(profile.excelPath.isEmpty() ? tr("(未设置excel)") : profile.excelPath);
    }

    m_lastTsDir = profile.tsDir;
    setTsFiles(tsFiles);

    m_bLoadingProfile = bLoading;

    QSettings settings(QSettings::IniFormat, QSettings::UserScope, "QtTranslateTool", "langImport");
    settings.setValue(kProfileNameKey, _name);

    if(missing.isEmpty())
    {
        return;
    }

    if(_interactive)
    {
        QMessageBox::warning(this, tr("匹配项文件缺失"),
                             tr("匹配项 \"%1\" 中以下文件不存在，已跳过：\n%2")
                             .arg(_name, missing.join("\n")));
    }
    else
    {
        statusBar()->showMessage(tr("匹配项 \"%1\" 中有 %2 个文件不存在").arg(_name).arg(missing.size()), 5000);
    }
}

void MainWindow::clearProfileSelection()
{
    if(m_curProfileName.isEmpty() && ui->combo_profile->currentIndex() == 0)
    {
        return;
    }

    m_curProfileName.clear();
    refreshProfileCombo();

    QSettings settings(QSettings::IniFormat, QSettings::UserScope, "QtTranslateTool", "langImport");
    settings.setValue(kProfileNameKey, QString());
}

QString MainWindow::currentProfileName() const
{
    if(ui->combo_profile->currentIndex() <= 0)
    {
        return QString();
    }
    return ui->combo_profile->currentText();
}

void MainWindow::showExcelError(const QString& _path)
{
    m_lastExcelPath.clear();
    getExcelOpt()->clearData();

    const QString name = _path.isEmpty()
            ? tr("未设置excel文件")
            : tr("%1 (文件不存在)").arg(QFileInfo(_path).fileName());
    m_pProgressExcel = initProgressBar(m_pProgressExcel, name);
    m_pProgressExcel->setError();
    ui->widget_excel->layout()->addWidget(m_pProgressExcel);
    ui->widget_excel->show();
}

void MainWindow::slot_profileChanged(int _index)
{
    if(m_bLoadingProfile)
    {
        return;
    }

    if(_index <= 0)
    {//回到“当前选择”，不再关联任何匹配项
        m_curProfileName.clear();
        QSettings settings(QSettings::IniFormat, QSettings::UserScope, "QtTranslateTool", "langImport");
        settings.setValue(kProfileNameKey, QString());
        return;
    }

    applyProfile(ui->combo_profile->itemText(_index), true);
}

void MainWindow::slot_saveProfile()
{
    if(m_lastExcelPath.isEmpty() && m_lststrPath.empty())
    {
        QMessageBox::information(this, tr("保存匹配项"), tr("请先选择excel和ts文件。"));
        return;
    }

    QString name = currentProfileName();
    if(name.isEmpty())
    {
        bool ok = false;
        name = QInputDialog::getText(this, tr("保存匹配项"), tr("匹配项名称："),
                                     QLineEdit::Normal, QString(), &ok).trimmed();
        if(!ok || name.isEmpty())
        {
            return;
        }

        if(m_mapProfiles.contains(name)
                && QMessageBox::Yes != QMessageBox::question(this, tr("保存匹配项"),
                                                             tr("匹配项 \"%1\" 已存在，是否覆盖？").arg(name)))
        {
            return;
        }
    }

    if(name.contains('/') || name.contains('\\'))
    {
        QMessageBox::warning(this, tr("保存匹配项"), tr("名称不能包含 / 或 \\ 字符。"));
        return;
    }

    MatchProfile profile;
    profile.excelPath = m_lastExcelPath;
    for(const QString& path : m_lststrPath)
    {
        profile.tsFiles.push_back(path);
    }
    profile.tsDir = lastTsDir();

    m_mapProfiles.insert(name, profile);
    saveProfiles();

    m_curProfileName = name;
    refreshProfileCombo(name);

    QSettings settings(QSettings::IniFormat, QSettings::UserScope, "QtTranslateTool", "langImport");
    settings.setValue(kProfileNameKey, name);
    statusBar()->showMessage(tr("匹配项 \"%1\" 已保存").arg(name), 3000);
}

void MainWindow::slot_deleteProfile()
{
    const QString name = currentProfileName();
    if(name.isEmpty())
    {
        QMessageBox::information(this, tr("删除匹配项"), tr("请先在下拉框中选择要删除的匹配项。"));
        return;
    }

    if(QMessageBox::Yes != QMessageBox::question(this, tr("删除匹配项"),
                                                 tr("确定删除匹配项 \"%1\" 吗？").arg(name)))
    {
        return;
    }

    m_mapProfiles.remove(name);
    saveProfiles();

    m_curProfileName.clear();
    refreshProfileCombo();

    QSettings settings(QSettings::IniFormat, QSettings::UserScope, "QtTranslateTool", "langImport");
    settings.setValue(kProfileNameKey, QString());
    statusBar()->showMessage(tr("匹配项 \"%1\" 已删除").arg(name), 3000);
}

QString MainWindow::lastTsDir() const
{
    if(!m_lastTsDir.isEmpty())
    {
        return m_lastTsDir;
    }
    if(!m_lststrPath.empty())
    {
        return QFileInfo(m_lststrPath.front()).absolutePath();
    }
    return QDir::homePath();
}

void MainWindow::handleDroppedFiles(const QStringList& _paths)
{
    QStringList excelFiles;
    QStringList tsFiles;

    for(const QString& path : _paths)
    {
        QFileInfo info(path);
        if(!info.isFile())
        {
            continue;
        }
        if(isExcelFile(path))
        {
            excelFiles.push_back(info.absoluteFilePath());
        }
        else if(isTsFile(path))
        {
            tsFiles.push_back(info.absoluteFilePath());
        }
    }

    if(!excelFiles.isEmpty())
    {
        loadExcelFile(excelFiles.first());
    }
    if(!tsFiles.isEmpty())
    {
        setTsFiles(tsFiles);
    }
}

void MainWindow::slot_filesDropped(const QStringList& _paths)
{
    handleDroppedFiles(_paths);
}

void MainWindow::dragEnterEvent(QDragEnterEvent* _event)
{
    if(_event->mimeData()->hasUrls())
    {
        for(const QUrl& url : _event->mimeData()->urls())
        {
            const QString path = url.toLocalFile();
            if(isExcelFile(path) || isTsFile(path))
            {
                _event->acceptProposedAction();
                return;
            }
        }
    }
    _event->ignore();
}

void MainWindow::dropEvent(QDropEvent* _event)
{
    QStringList paths;
    if(_event->mimeData()->hasUrls())
    {
        for(const QUrl& url : _event->mimeData()->urls())
        {
            const QString path = url.toLocalFile();
            if(!path.isEmpty())
            {
                paths.push_back(path);
            }
        }
    }

    if(paths.isEmpty())
    {
        _event->ignore();
        return;
    }

    _event->acceptProposedAction();
    handleDroppedFiles(paths);
}
