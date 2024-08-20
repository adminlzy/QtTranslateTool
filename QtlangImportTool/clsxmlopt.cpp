#include "clsxmlopt.h"
#include <QDebug>
#include <QFile>
#include <QDomDocument>
#include <QTextStream>
#include <QThread>
#include <QFileInfo>

clsXmlOpt::clsXmlOpt(QObject *parent) : QObject(parent)
{
    connect(this, &clsXmlOpt::signal_finishImport, this, &clsXmlOpt::slot_finishImport);
}

int clsXmlOpt::addWordData(const QString& _file, const QMap<wordId, trans>& _mapExcelData)
{
    int ikey = newKey();
    std::thread* pthr = new std::thread(&clsXmlOpt::thrImportExcelData, this, ikey, _file, _mapExcelData);
    m_mapthrImport.insert(ikey, pthr);
    return ikey;
}

int clsXmlOpt::addLoadTsTask(const QString& _file)
{
    int ikey = newKey();
    std::thread* pthr = new std::thread(&clsXmlOpt::loadTs, this, ikey, _file);
    m_mapthrImport.insert(ikey, pthr);
    return ikey;
}

void clsXmlOpt::thrImportExcelData(int _key, QString _file, QMap<wordId, trans> _mapExcelData)
{
    qDebug()<<"clsXmlOpt::thrImportExcelData _key:"<<_key<<" _file:"<<_file;
    for (const auto &pair : _mapExcelData) {
        qDebug()<<pair;
    }

    QFile file(_file);
    if (!file.open(QFileDevice::ReadOnly)) {
        qDebug()<<"文件打开失败！";
        return;
    }

    QDomDocument doc;
    if (!doc.setContent(&file)) {
        qDebug()<<"操作的文件不是XML文件！";
        file.close();
        return;
    }
    file.close();

    QDomElement root = doc.documentElement();
    qDebug() << "根节点：" << root.nodeName();

    // 获取所有Book1节点
    QDomNodeList list = root.elementsByTagName("message");

    /* 获取属性中的值 */
    for (int i = 0; i < list.count(); i++)
    {
        emit signal_importTsProcess(_key, (100/list.count())*i, false);
        // 获取链表中的值
        QDomElement element = list.at(i).toElement();
        QDomElement source = element.firstChildElement("source");
        QString srcWord = source.text();
        qDebug()<<"srcWord:"<<srcWord;
        auto itor = _mapExcelData.find(srcWord);
        if(itor == _mapExcelData.end())
        {
            itor = _mapExcelData.find(srcWord.trimmed());
            if(itor == _mapExcelData.end())
            {
                qDebug()<<"_mapExcelData not find";
                continue;
            }

        }

        QString result = itor.value();
        QDomNodeList sublist = element.childNodes();
        for(int j = 0; j < sublist.count(); j++)
        {
            QDomNode subnode = sublist.at(j);
            if("translation" != subnode.toElement().tagName())
            {
                qDebug()<<"translation != subnode.toElement().tagName()";
                continue;
            }
#if 1//覆盖
            if(!m_replace){
                if(!subnode.toElement().text().isEmpty())
                {
                    qDebug()<<"!subnode.toElement().text().isEmpty()";
                    continue;
                }
            }
#endif
            qDebug()<<"trans:"<<result;
            QDomElement translation = doc.createElement("translation");
            if(!result.isEmpty())
            {
                translation.appendChild(doc.createTextNode(result));
                element.replaceChild(translation, subnode);
            }
        }

    }

    if (!file.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        // 处理文件打开失败的情况
        return;
    }
    QTextStream out(&file);
    out.setCodec("UTF-8");
    out << doc.toString();
    file.close();

    emit signal_importTsProcess(_key, 100, true);
    emit signal_finishImport(_key);//回收线程
}

void clsXmlOpt::slot_finishImport(int _key)
{
    auto itor = m_mapthrImport.find(_key);
    if(itor == m_mapthrImport.end())
    {
        return;
    }
    std::thread* pthr = itor.value();

    if(pthr)
    {
        if(pthr->joinable())
        {
            pthr->join();
        }
    }
    m_mapthrImport.erase(itor);
}

void clsXmlOpt::loadTs(int _key, const QString &_file)
{
    m_mapWordData.clear();
    m_mapEmptyWordData.clear();
    qDebug()<<"clsXmlOpt::loadTs _file:"<<_file;
    QFileInfo ts(_file);
    QString lang = ts.baseName().right(2);
    m_setLangs.insert(lang);
    QFile file(_file);
    if (!file.open(QFileDevice::ReadOnly)) {
        qDebug()<<"文件打开失败！";
        return;
    }

    QDomDocument doc;
    if (!doc.setContent(&file)) {
        qDebug()<<"操作的文件不是XML文件！";
        file.close();
        return;
    }
    file.close();

    QDomElement root = doc.documentElement();
    qDebug() << "根节点：" << root.nodeName();

    // 获取所有Book1节点
    QDomNodeList list = root.elementsByTagName("message");

    /* 获取属性中的值 */
    for (int i = 0; i < list.count(); i++)
    {
        emit signal_loadTsProcess(_key, (100*i/list.count()), false);
        // 获取链表中的值
        QDomElement element = list.at(i).toElement();
        QDomElement source = element.firstChildElement("source");
        QDomElement translation = element.firstChildElement("translation");

        QString typeAttr = translation.attribute("type");
        qDebug() << "Attribute value:" << typeAttr;
        if("vanished" == typeAttr
                || "obsolete" == typeAttr){
            //已经不使用了
            continue;
        }
        QString srcWord = source.text();
        QString transWord = translation.text();
        insertTrans(m_mapWordData, srcWord, lang, transWord);
        if(transWord.isEmpty()){
            insertTrans(m_mapEmptyWordData, srcWord, lang, transWord);
        }
    }

    emit signal_loadTsProcess(_key, 100, true);

    emit signal_finishImport(_key);//回收线程
}

void clsXmlOpt::insertTrans(QMap<wordId, QMap<langId, trans>>& _mapWordData, const QString& _wordId, const QString& _langId, const QString& _trans)
{
    std::lock_guard<std::mutex> g(mtxTrans);

    auto itor = _mapWordData.find(_wordId);
    if(itor == _mapWordData.end()){
        QMap<langId, trans> mapTrans;
        mapTrans.insert(_langId, _trans);
        _mapWordData.insert(_wordId, mapTrans);
        return;
    }

    auto& mapTrans = itor.value();
    auto itorSub = mapTrans.find(_langId);
    if(itorSub == mapTrans.end()){
        mapTrans.insert(_langId, _trans);
    }
}

QMap<wordId, QMap<langId, trans>> clsXmlOpt::getWordData()
{
    return m_mapWordData;
}

QMap<wordId, QMap<langId, trans>> clsXmlOpt::getEmptyWordData()
{
    return m_mapEmptyWordData;
}

QSet<langId> clsXmlOpt::getLangs()
{
    return m_setLangs;
}
