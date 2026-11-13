#include "dropdownarea.h"
#include <QDragEnterEvent>
#include <QDragLeaveEvent>
#include <QDropEvent>
#include <QFileInfo>
#include <QMimeData>
#include <QUrl>

namespace {
bool isSupportedSuffix(const QString& _suffix)
{
    const QString suffix = _suffix.toLower();
    return suffix == "xls" || suffix == "xlsx" || suffix == "ts";
}
}

DropArea::DropArea(QWidget *parent) : QLabel(parent)
{
    setAcceptDrops(true);
    setAlignment(Qt::AlignCenter);
    setText(tr("将 Excel（.xls / .xlsx）或 TS（.ts）文件拖拽到此处"));
    setMinimumHeight(84);

    m_strNormalStyle = "QLabel {"
                       "border: 2px dashed #c3d0e4;"
                       "border-radius: 8px;"
                       "background-color: #f8fafd;"
                       "color: #5b6b82;"
                       "font-size: 13px;"
                       "}";
    m_strActiveStyle = "QLabel {"
                       "border: 2px dashed #2f6fed;"
                       "border-radius: 8px;"
                       "background-color: #eef4ff;"
                       "color: #1d5bd8;"
                       "font-size: 13px;"
                       "}";
    setStyleSheet(m_strNormalStyle);
}

void DropArea::setActive(bool _active)
{
    setStyleSheet(_active ? m_strActiveStyle : m_strNormalStyle);
}

bool DropArea::containsSupportedFile(const QMimeData* _mimeData) const
{
    if(!_mimeData->hasUrls())
    {
        return false;
    }

    for(const QUrl& url : _mimeData->urls())
    {
        QFileInfo info(url.toLocalFile());
        if(info.isFile() && isSupportedSuffix(info.suffix()))
        {
            return true;
        }
    }
    return false;
}

void DropArea::dragEnterEvent(QDragEnterEvent* _event)
{
    if(containsSupportedFile(_event->mimeData()))
    {
        setActive(true);
        _event->acceptProposedAction();
        return;
    }
    _event->ignore();
}

void DropArea::dragLeaveEvent(QDragLeaveEvent* _event)
{
    setActive(false);
    _event->accept();
}

void DropArea::dropEvent(QDropEvent* _event)
{
    setActive(false);

    QStringList paths;
    if(_event->mimeData()->hasUrls())
    {
        for(const QUrl& url : _event->mimeData()->urls())
        {
            QString path = url.toLocalFile();
            QFileInfo info(path);
            if(info.isFile() && isSupportedSuffix(info.suffix()))
            {
                paths.push_back(info.absoluteFilePath());
            }
        }
    }

    if(paths.isEmpty())
    {
        _event->ignore();
        return;
    }

    _event->acceptProposedAction();
    emit filesDropped(paths);
}
