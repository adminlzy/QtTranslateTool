#include "clsexcelopt.h"
#include <QAxObject>
#include <QVariant>
#include <QDebug>
#include <vector>
#include <Windows.h>
#include <QThread>
#include <QDir>

//版权声明：本文为博主原创文章，遵循 CC 4.0 BY-SA 版权协议，转载请附上原文出处链接和本声明。
//原文链接：https://blog.csdn.net/weixin_43519792/article/details/103431223

clsExcelOpt::clsExcelOpt(QObject *parent) : QObject(parent)
{
    connect(this, &clsExcelOpt::signal_readExcel, this, &clsExcelOpt::slot_readExcel, Qt::QueuedConnection);
    connect(this, &clsExcelOpt::signal_fillExcel, this, &clsExcelOpt::slot_fillExcel, Qt::QueuedConnection);
    QThread* pthr = new QThread(this);
    pthr->start();
    this->moveToThread(pthr);
    qDebug()<<"clsExcelOpt::clsExcelOpt into";
}

QString clsExcelOpt::getCurPath()
{
    return m_curPath;
}

void clsExcelOpt::readCurSheet(const QString& _path)
{
    m_curPath = _path;
    qDebug()<<"clsExcelOpt::readCurSheet into";
    emit signal_readExcel(_path);
}

void clsExcelOpt::slot_readExcel(QString _path)
{
    qDebug()<<"clsExcelOpt::slot_readExcel into";
    readExcel(_path);
}

void clsExcelOpt::readExcel(QString _path)
{
    m_mapExcelData.clear();
    m_setWordId.clear();

    CoInitializeEx(NULL, COINIT_MULTITHREADED);
    QAxObject* pexcel = nullptr;                            //操作Excel文件对象(open-save-close-quit)
    QAxObject* pworkbooks = nullptr;                        //总工作薄对象
    QAxObject* pworkbook = nullptr;                         //操作当前工作薄对象
    QAxObject* pworksheets = nullptr;                       //文件中所有<Sheet>表页
    QAxObject* pworksheet = nullptr;                        //存储第n个sheet对象
    QAxObject* pusedrange = nullptr;                        //存储当前sheet的数据对象

    pexcel = new QAxObject("Excel.Application");									//创建Excel对象连接驱动
    pexcel->dynamicCall("SetVisible(bool)",false);								//ture的打开Excel表 false不打开Excel表
    pexcel->setProperty("DisplayAlerts",false);
    pworkbooks = pexcel->querySubObject("WorkBooks");
    pworkbook = pworkbooks->querySubObject("Open(const QString&)", _path);         //打开指定Excel
    pworksheets = pworkbook->querySubObject("WorkSheets");            			//获取表页对象
    pworksheet = pworksheets->querySubObject("Item(int)",1);          			//获取第1个sheet表
    pusedrange = pworksheet->querySubObject("Usedrange");							//获取权限
    int intRow = pusedrange->querySubObject("Rows")->property("Count").toInt();  //获取数据总行数
    int intCol = pusedrange->querySubObject("Columns")->property("Count").toInt();  //获取数据总行数

    qDebug()<<"intRow:"<<intRow<<"intCol:"<<intCol;

    //先找出所有列
    std::vector<wordId> vecRowKey;
    for(int i = 2; i <= intRow; i++)
    {
        wordId srcWord = getCell(pworksheet, i, 1);
        if(srcWord.isEmpty())
        {
            qDebug()<<"intRow key error null i:"<<i;
        }
        qDebug()<<"excel srcWord:"<<srcWord;
        m_setWordId.insert(srcWord);
        vecRowKey.push_back(srcWord);

        int iprocess = (100*(intRow + i))/(intRow * intCol);
        emit signalReadProcess(iprocess, false);
    }

    for(int j = 2; j <= intCol; j++)
    {
        QString lang = getCell(pworksheet, 1, j);
        if(lang.isEmpty())
        {
            break;
        }

        lang = getSmallLang(lang);
        if(lang.isEmpty()){
            continue;
        }
        qDebug()<<lang;
        m_ExcelLangs.insert(lang);

        QMap<wordId, trans> mapCol;
        //逐行读取主表
        for (int i = 2; i <= intRow; i++)
        {
            trans value = getCell(pworksheet, i, j);
            if(value.isEmpty())
            {
                continue;
            }

            if(!vecRowKey[i - 2].isEmpty())
            {
                mapCol[vecRowKey[i - 2]] = value;
            }

            int iprocess = (100*(intRow*j + i))/(intRow * intCol);
            emit signalReadProcess(iprocess, false);
        }

        m_mapExcelData.insert(lang, mapCol);
    }

    pworkbook->dynamicCall("Close(Boolean)",false);
    pexcel->dynamicCall("Quit(void)");

    emit signalReadProcess(100, true);
}

QMap<wordId, trans> clsExcelOpt::getExcelData(langId _key)
{
    auto itor = m_mapExcelData.find(_key);
    if(itor == m_mapExcelData.end())
    {
        return QMap<wordId, trans>();
    }
    return itor.value();
}

void clsExcelOpt::addTsData(const QMap<wordId, QMap<langId, trans>>& tsData, QMap<wordId, QMap<langId, trans>>& tsEmptyData, QSet<langId>& _langs)
{
    m_tsData = tsData;
    m_tsEmptyData = tsEmptyData;
    m_tsLangs = _langs;
    emit signal_fillExcel(m_curPath);
}
void clsExcelOpt::createNoTransFile(QString _path, int _totalPro)
{
    // 创建Excel应用对象
    QAxObject *excel = new QAxObject("Excel.Application");
    // 设置是否显示Excel窗口
    excel->setProperty("Visible", false);

    // 创建一个新的工作簿
    QAxObject *workbooks = excel->querySubObject("Workbooks");
    QAxObject *workbook = workbooks->querySubObject("Add()");
    // 获取第一个工作表
    QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1);

    int i = 2;
    int colMax = 0;
    for(auto itor = m_tsEmptyData.begin() ; itor != m_tsEmptyData.end(); itor++, i++){
        QString srcWord = itor.key();
        auto itorFull = m_tsData.find(srcWord);
        setCell(worksheet, i, 1, itorFull.key());

        QMap<langId, trans> mapTrans = itorFull.value();
        for(auto itorSub = mapTrans.begin(); itorSub != mapTrans.end(); itorSub++){
            langId lang = itorSub.key();
            lang = getFullLang(lang);
            int editCol = 2;
            while(editCol <= colMax){ //查找是否已有语言表头
                QString langHead = getCell(worksheet, 1, editCol);
                if(lang == langHead){
                    break;
                }
                editCol++;
            }
            if(editCol > colMax){ //未找到语言表头 添加一列语言
                colMax = editCol;
                setCell(worksheet, 1, colMax, lang);
                qDebug()<<QString("head:%1 row:%2 col:%3").arg(lang).arg(1).arg(colMax);
            }

            setCell(worksheet, i, editCol, itorSub.value());

        }
        emit signalReadProcess(_totalPro*i/m_tsEmptyData.size(), false);
    }

    // 保存工作簿
    workbook->dynamicCall("SaveAs(const QString&)", _path);

    // 关闭工作簿
    workbook->dynamicCall("Close(bool)", false);

    // 退出Excel应用
    excel->dynamicCall("Quit()");

    // 释放对象
    delete worksheet;
    delete workbook;
    delete workbooks;
    delete excel;
}
void clsExcelOpt::slot_fillExcel(QString _path)
{
    if(_path.isEmpty()){
        return;
    }
    if(m_tsData.isEmpty()){
        return;
    }


    //先将为空的ts数据补齐
    for(auto itor = m_tsEmptyData.begin(); itor != m_tsEmptyData.end(); itor++){
        QString strWordId = itor.key();
        QMap<langId, trans> wordData = itor.value();

        auto itorTs = m_tsData.find(strWordId);
        auto& mapTsWord = itorTs.value();

        for(auto itorSub = wordData.begin(); itorSub != wordData.end(); ){
            QString lang = itorSub.key();
            QMap<wordId, trans> mapExcelWord;
            auto itorExcel = m_mapExcelData.find(lang);
            if(itorExcel != m_mapExcelData.end()){
                mapExcelWord = itorExcel.value();
            }

            auto itorExeclSub = mapExcelWord.find(strWordId);
            if(itorExeclSub != mapExcelWord.end()){ //word中找到词条
                mapTsWord[lang] = itorExeclSub.value(); //将excel中的翻译补到ts容器中
                itorSub = wordData.erase(itorSub); //词条有翻译 清理
                continue;
            }
            itorSub++;
        }
    }

    //todo 将没翻译的词条写到excel中
    QString workbookPath = QDir::currentPath() + "/noTransWord.xlsx";
    createNoTransFile(workbookPath, 50);


    //todo 将完整的ts写到 excel中
    CoInitializeEx(NULL, COINIT_MULTITHREADED);
    QAxObject* pexcel = nullptr;                            //操作Excel文件对象(open-save-close-quit)
    QAxObject* pworkbooks = nullptr;                        //总工作薄对象
    QAxObject* pworkbook = nullptr;                         //操作当前工作薄对象
    QAxObject* pworksheets = nullptr;                       //文件中所有<Sheet>表页
    QAxObject* pworksheet = nullptr;                        //存储第n个sheet对象
    QAxObject* pusedrange = nullptr;                        //存储当前sheet的数据对象

    pexcel = new QAxObject("Excel.Application");									//创建Excel对象连接驱动
    pexcel->dynamicCall("SetVisible(bool)",false);								//ture的打开Excel表 false不打开Excel表
    pexcel->setProperty("DisplayAlerts",false);
    pworkbooks = pexcel->querySubObject("WorkBooks");
    pworkbook = pworkbooks->querySubObject("Open(const QString&)", _path);         //打开指定Excel
    pworksheets = pworkbook->querySubObject("WorkSheets");            			//获取表页对象
    pworksheet = pworksheets->querySubObject("Item(int)",1);          			//获取第1个sheet表
    pusedrange = pworksheet->querySubObject("Usedrange");							//获取权限
    int intRow = pusedrange->querySubObject("Rows")->property("Count").toInt();  //获取数据总行数
    int intCol = pusedrange->querySubObject("Columns")->property("Count").toInt();  //获取数据总行数

    qDebug()<<"intRow:"<<intRow<<"intCol:"<<intCol;

    //补充新词条
    for(auto itor = m_tsData.begin(); itor != m_tsData.end(); itor++){
        wordId srcWord = itor.key();
        auto itorWordId = m_setWordId.find(srcWord);
        if(itorWordId != m_setWordId.end()){ //已有词条
            continue;
        }
        qDebug()<<"new add wordId:"<<srcWord;
        setCell(pworksheet, ++intRow, 1, srcWord);
    }

    //补充新语言
    for(langId lang : m_tsLangs){
        auto itor = m_ExcelLangs.find(lang);
        if(itor != m_ExcelLangs.end()){
            continue;
        }
        lang = getFullLang(lang);
        qDebug()<<"new add langId:"<<lang;
        setCell(pworksheet, 1, ++intCol, lang);
    }

    //补充excel翻译为空的单元格
    for (int i = 2; i <= intRow; i++){
        wordId srcWord = getCell(pworksheet, i, 1);
        qDebug()<<"srcWord:"<<srcWord;
        auto itorTs = m_tsData.find(srcWord);
        if(itorTs == m_tsData.end()){
            continue;
        }

        auto mapTrans = itorTs.value();
        for(int j = 2; j <= intCol; j++){
            trans srcTrans = getCell(pworksheet, i, j);
#if 1 //覆盖
            if(!m_replace){
                if(!srcTrans.isEmpty()){//excel中翻译为空的单元格
                    continue;
                }
            }

#endif
            qDebug()<<"not have trans:"<<srcWord;
            QString langHead = getCell(pworksheet, 1, j);
            langHead = getSmallLang(langHead);
            if(langHead.isEmpty()){
                continue;
            }
            qDebug().noquote()<<" "<<langHead;
            auto itor = mapTrans.find(langHead);
            if(itor == mapTrans.end()){//查找ts中的翻译
                continue;
            }

            QString transTs = itor.value();
            if(transTs.isEmpty()){
                continue;
            }
            qDebug().noquote()<<" transTs:"<<transTs;
            if(!transTs.isEmpty()){//将ts文件中的翻译填充到 excel单元格
                setCell(pworksheet, i, j, transTs);
            }

            int iprocess = (50*(intCol*i + j))/(intRow * intCol);
            emit signalReadProcess(iprocess + 50, false);
        }
    }



    // 保存工作簿
#if 0
    QString newWorkbookPath = QDir::currentPath() + "/newTransWord.xlsx";
    pworkbook->dynamicCall("SaveAs(const QString&)", newWorkbookPath);
#else
    pworkbook->dynamicCall("Save()");
#endif

    pworkbook->dynamicCall("Close(Boolean)",false);
    pexcel->dynamicCall("Quit(void)");

    delete pusedrange;
    delete pworksheet;
    delete pworksheets;
    delete pworkbook;
    delete pworkbooks;
    delete pexcel;

    emit signalReadProcess(100, true);
}

QString clsExcelOpt::getCell(QAxObject *_pworksheet, int _row, int _col)
{
    QAxObject* cellHead = _pworksheet->querySubObject("Cells(int,int)", _row, _col);
    QString valueCell = cellHead->dynamicCall(("Value2()")).value<QString>();
    delete cellHead;
    return valueCell;
}

void clsExcelOpt::setCell(QAxObject *_pworksheet, int _row, int _col, QString _value)
{
    QAxObject *cell = _pworksheet->querySubObject("Cells(int,int)", _row, _col);
    cell->dynamicCall("Value", _value);
    delete cell;
}

QString clsExcelOpt::getFullLang(langId _lang)
{
    if("EN" == _lang){
        return "English";
    }
    else if("SP" == _lang){
        return "Spanish";
    }
    else if("FR" == _lang){
        return "French";
    }
    else if("PO" == _lang){
        return "Portuguese";
    }
    else if("RU" == _lang){
        return "Russian";
    }
    else if("VI" == _lang){
        return "Vietnamese";
    }
    else if("UK" == _lang){
        return "Ukrainian";
    }
    else if("IT" == _lang){
        return "Italian";
    }
    else if("TU" == _lang){
        return "Turkish";
    }
    return "";
}

langId clsExcelOpt::getSmallLang(QString _lang){
    if("English" == _lang){
        return "EN";
    }
    else if("Spanish" == _lang){
        return "SP";
    }
    else if("French" == _lang){
        return "FR";
    }
    else if("Portuguese" == _lang){
        return "PO";
    }
    else if("Russian" == _lang){
        return "RU";
    }
    else if("Vietnamese" == _lang){
        return "VI";
    }
    else if("Ukrainian" == _lang){
        return "UK";
    }
    else if("Italian" == _lang){
        return "IT";
    }
    else if("Turkish" == _lang){
        return "TU";
    }
    return "";
}
