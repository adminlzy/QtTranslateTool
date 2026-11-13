#ifndef DROPAREA_H
#define DROPAREA_H

#include <QLabel>
#include <QStringList>

class QMimeData;

class DropArea : public QLabel
{
    Q_OBJECT
public:
    explicit DropArea(QWidget *parent = nullptr);

signals:
    void filesDropped(const QStringList& _paths);

protected:
    void dragEnterEvent(QDragEnterEvent* _event) override;
    void dragLeaveEvent(QDragLeaveEvent* _event) override;
    void dropEvent(QDropEvent* _event) override;

private:
    bool containsSupportedFile(const QMimeData* _mimeData) const;
    void setActive(bool _active);

    QString m_strNormalStyle;
    QString m_strActiveStyle;
};

#endif // DROPAREA_H
