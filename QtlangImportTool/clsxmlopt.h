#ifndef CLSXMLOPT_H
#define CLSXMLOPT_H

#include <QObject>
#include "clsexcelopt.h"
#include <thread>
#include <QMap>
#include <mutex>
#include <QSet>

class clsXmlOpt : public QObject
{
    Q_OBJECT
public:
    explicit clsXmlOpt(QObject *parent = nullptr);

    int addWordData(const QString& _file, const QMap<wordId, trans>& _mapExcelData);
    int addLoadTsTask(const QString& _file);
    int newKey()
    {
        return m_iKey++;
    }
    QMap<wordId, QMap<langId, trans>> getWordData();
    QMap<wordId, QMap<langId, trans>> getEmptyWordData();

    QSet<langId> getLangs();

    void setReplace(bool _replace){
        m_replace = _replace;
    }
signals:
    void signal_importTsProcess(int key, int _process, bool _finish);
    void signal_loadTsProcess(int key, int _process, bool _finish);
    void signal_finishImport(int _key);

private slots:
    void slot_finishImport(int _key);
private:

    void thrImportExcelData(int _key, QString _file, QMap<wordId, trans> _mapExcelData);

    /*
     * 读取ts文件
     */
    void loadTs(int _key, const QString& _file);

    void insertTrans(QMap<wordId, QMap<langId, trans> > &_mapWordData, const QString& _wordId, const QString& _langId, const QString& _trans);
private:
    QMap<int, std::thread*> m_mapthrImport;
    int m_iKey = 0;

    QSet<langId> m_setLangs;
    QMap<wordId, QMap<langId, trans>> m_mapWordData;//读取所有的ts词条
    QMap<wordId, QMap<langId, trans>> m_mapEmptyWordData;//读取没有翻译的词条
    std::mutex mtxTrans;
    bool m_replace = false;
};

#endif // CLSXMLOPT_H
