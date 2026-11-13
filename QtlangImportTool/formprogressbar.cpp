#include "formprogressbar.h"
#include "ui_formprogressbar.h"
#include <QFileInfo>
#include <QResizeEvent>
#include <QStyle>
#include <QTimer>

FormProgressBar::FormProgressBar(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::FormProgressBar)
{
    ui->setupUi(this);
    ui->progressBar->setRange(0, 100);
    ui->progressBar->setValue(0);
    ui->label_percent->setMinimumWidth(40);
    ui->label_percent->setAlignment(Qt::AlignRight | Qt::AlignVCenter);
    m_strState.clear();
    setState(QStringLiteral("normal"));
}

FormProgressBar::~FormProgressBar()
{
    delete ui;
}

void FormProgressBar::setName(const QString& _qstrName)
{
    m_strName = _qstrName;
    const QString fileName = QFileInfo(_qstrName).fileName();
    m_strDisplayName = fileName.isEmpty() ? _qstrName : fileName;
    ui->label_name->setToolTip(_qstrName);
    updateElidedName();
}

void FormProgressBar::setProgress(int _pro)
{
    const int value = qBound(0, _pro, 100);
    ui->progressBar->setValue(value);

    if(m_strState == QLatin1String("error"))
    {
        ui->label_percent->setText(tr("失败"));
        return;
    }

    ui->label_percent->setText(QString::number(value) + "%");
    ui->label_percent->setToolTip(QString());
    setState(value >= 100 ? QStringLiteral("done") : QStringLiteral("normal"));
}

void FormProgressBar::setError()
{
    setState(QStringLiteral("error"));
    ui->progressBar->setValue(100);
    ui->label_percent->setText(tr("失败"));
    ui->label_percent->setToolTip(tr("文件不存在或读取失败"));
}

void FormProgressBar::setState(const QString& _state)
{
    if(m_strState == _state)
    {
        return;
    }

    m_strState = _state;
    ui->progressBar->setProperty("state", _state);
    ui->label_percent->setProperty("state", _state);
    refreshStyle(ui->progressBar);
    refreshStyle(ui->label_percent);
}

void FormProgressBar::refreshStyle(QWidget* _widget)
{
    _widget->style()->unpolish(_widget);
    _widget->style()->polish(_widget);
    _widget->update();
}

void FormProgressBar::updateElidedName()
{
    const int width = ui->label_name->width();
    if(width <= 0 || m_strDisplayName.isEmpty())
    {
        ui->label_name->setText(m_strDisplayName);
        return;
    }
    ui->label_name->setText(ui->label_name->fontMetrics().elidedText(m_strDisplayName,
                                                                    Qt::ElideMiddle, width));
}

void FormProgressBar::resizeEvent(QResizeEvent* _event)
{
    QWidget::resizeEvent(_event);
    QTimer::singleShot(0, this, &FormProgressBar::updateElidedName);
}
