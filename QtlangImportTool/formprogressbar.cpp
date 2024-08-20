#include "formprogressbar.h"
#include "ui_formprogressbar.h"
#include <QDebug>

FormProgressBar::FormProgressBar(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::FormProgressBar)
{
    ui->setupUi(this);
    ui->progressBar->setRange(0, 100);
    ui->progressBar->setValue(0);
    ui->progressBar->setFormat("%v");
    QString s1 = "QProgressBar {\
        border: 2px solid grey;\
        border-radius: 5px;\
        text-align: center;\
        color:#006400;\
        font-weight: bold;\
    }";

    QString s2 = "QProgressBar::chunk {\
        background-color: #42b983;\
        width: 20px;\
        margin: 0.5px;\
    }";
    ui->progressBar->setStyleSheet(s1 + s2);
    this->setMaximumHeight(50);
}

FormProgressBar::~FormProgressBar()
{
    delete ui;
}

void FormProgressBar::setName(const QString& _qstrName)
{
    m_strName = _qstrName;
}

void FormProgressBar::setProgress(int _pro)
{
    ui->progressBar->setValue(_pro);
    ui->progressBar->setFormat(QString("%1:%2%").arg(m_strName).arg(_pro));
}

void FormProgressBar::setError()
{
    QString s1 = "QProgressBar {\
        border: 2px solid grey;\
        border-radius: 5px;\
        text-align: center;\
        color:#800000;\
        font-weight: bold;\
    }";

    QString s2 = "QProgressBar::chunk {\
        background-color: #FF0000;\
        width: 20px;\
        margin: 0.5px;\
    }";
    ui->progressBar->setStyleSheet(s1 + s2);
}
