#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <list>
#include <QMap>
#include <QStringList>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class clsExcelOpt;
class clsXmlOpt;
class FormProgressBar;

struct MatchProfile
{
    QString excelPath;
    QStringList tsFiles;
    QString tsDir;
};

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

protected:
    void dragEnterEvent(QDragEnterEvent* _event) override;
    void dropEvent(QDropEvent* _event) override;

private slots:

    void slot_readExcelProcess(int _process, bool _finish);
    void slot_importTsProcess(int key, int _process, bool _finish);
    void slot_loadTsProcess(int key, int _process, bool _finish);
    void on_pushButton_loadexcel_clicked();

    void on_pushButton_selectTsfile_clicked();

    void on_btn_excelts_clicked();

    void on_btn_tsexcel_clicked();

    void on_pushButton_reload_clicked();

    void slot_filesDropped(const QStringList& _paths);

    void slot_profileChanged(int _index);
    void slot_saveProfile();
    void slot_deleteProfile();

private:
    void readExcel();
    void loadExcelFile(const QString& _fileName);
    void setTsFiles(const QStringList& _fileNames);
    void handleDroppedFiles(const QStringList& _paths);
    void restoreLastSelection();
    void saveSettings();
    QString lastTsDir() const;

    void loadProfiles();
    void saveProfiles();
    void refreshProfileCombo(const QString& _selected = QString());
    void applyProfile(const QString& _name, bool _interactive);
    void clearProfileSelection();
    QString currentProfileName() const;
    void showExcelError(const QString& _path);

    void starfillTs();
    void starfillExcel();
    void loadTs();

    FormProgressBar* initProgressBar(FormProgressBar* _Bar, const QString& _name, bool _blerr=false);

    clsXmlOpt* getXmlOpt();
    clsExcelOpt* getExcelOpt();

    void createbtn(int);
private:
    Ui::MainWindow *ui;
    clsExcelOpt* m_pexcel = nullptr;
    clsXmlOpt* m_xmlTx = nullptr;
    std::list<QString> m_lststrPath;
    QMap<int, FormProgressBar*> m_mapTask;
    FormProgressBar* m_pProgressExcel = nullptr; //读取excel文件进度
    QString m_lastExcelPath; //上次选择的excel文件
    QString m_lastTsDir; //上次选择ts文件所在目录

    QMap<QString, MatchProfile> m_mapProfiles; //自定义匹配项
    QString m_curProfileName; //当前选中的匹配项
    bool m_bLoadingProfile = false; //正在按匹配项加载，避免回写选择
};
#endif // MAINWINDOW_H
