#ifndef FORMPROGRESSBAR_H
#define FORMPROGRESSBAR_H

#include <QWidget>

namespace Ui {
class FormProgressBar;
}

class FormProgressBar : public QWidget
{
    Q_OBJECT

public:
    explicit FormProgressBar(QWidget *parent = nullptr);
    ~FormProgressBar();
    void setError();
    void setName(const QString& _qstrName);
    void setProgress(int _pro);
private:
    Ui::FormProgressBar *ui;
    QString m_strName;
};

#endif // FORMPROGRESSBAR_H
