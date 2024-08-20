#ifndef CLSEXCELOPT_H
#define CLSEXCELOPT_H

#include <QObject>
#include <QMap>
#include <QSet>

typedef QString wordId;
typedef QString langId;
typedef QString trans;
class QAxObject;
class clsExcelOpt : public QObject
{
    Q_OBJECT
public:
    explicit clsExcelOpt(QObject *parent = nullptr);

    void readCurSheet(const QString& _path);

    QMap<wordId, trans> getExcelData(langId _key);

    void addTsData(const QMap<wordId, QMap<langId, trans>>&, QMap<wordId, QMap<langId, trans>>&, QSet<langId> &_langs);
    QString getCurPath();

    void setReplace(bool _replace){
        m_replace = _replace;
    }
signals:
    void signalReadProcess(int _process, bool _finish);
    void signal_readExcel(QString _path);

    void signal_fillExcel(QString _path);
private slots:
    void slot_readExcel(QString _path);
    void slot_fillExcel(QString _path);
private:
    void readExcel(QString _path);

    void createNoTransFile(QString _path, int _totalPro);
    /*
     * 单元格设置获取操作
     */
    QString getCell(QAxObject *_pworksheet, int _row, int _col);
    void setCell(QAxObject *_pworksheet, int _row, int _col, QString _value);

    QString getFullLang(langId _lang);
    langId getSmallLang(QString _lang);
private:
    QString m_curPath = ""; //当前加载的excel路径
    QSet<wordId> m_setWordId;
    QMap<langId, QMap<wordId, trans>> m_mapExcelData;

    QMap<wordId, QMap<langId, trans>> m_tsData; //ts文件完整数据
    QMap<wordId, QMap<langId, trans>> m_tsEmptyData; //ts文件词条为空数据
    QSet<langId> m_tsLangs;
    QSet<langId> m_ExcelLangs;
    bool m_replace = false;
};

#endif // CLSEXCELOPT_H
