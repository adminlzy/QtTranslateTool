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

protected:
    void resizeEvent(QResizeEvent* _event) override;

private:
    void setState(const QString& _state);
    void refreshStyle(QWidget* _widget);
    void updateElidedName();

    Ui::FormProgressBar *ui;
    QString m_strName;
    QString m_strDisplayName;
    QString m_strState = QStringLiteral("normal");
};

#endif // FORMPROGRESSBAR_H
