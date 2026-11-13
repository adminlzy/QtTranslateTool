#include "mainwindow.h"
#include <QApplication>
#include <QFile>
#include <QIcon>

int main(int argc, char *argv[])
{
    QApplication::setAttribute(Qt::AA_EnableHighDpiScaling);
    QApplication::setAttribute(Qt::AA_UseHighDpiPixmaps);

    QApplication a(argc, argv);
    a.setWindowIcon(QIcon(QStringLiteral(":/icons/app.svg")));

    QFile qssFile(QStringLiteral(":/style.qss"));
    if(qssFile.open(QIODevice::ReadOnly | QIODevice::Text))
    {
        a.setStyleSheet(QString::fromUtf8(qssFile.readAll()));
    }

    MainWindow w;
    w.show();
    return a.exec();
}
