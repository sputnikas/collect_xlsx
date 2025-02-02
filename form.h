#pragma once

#include <QMainWindow>
#include <QFileDialog>
#include <QCheckBox>
#include <QVBoxLayout>
#include <QMessageBox>
#include "ui_form.h"
#include "collect_xlsx.h"

class form : public QMainWindow
{
    Q_OBJECT
   public:
    form(QWidget *parent = nullptr);
    ~form();

    void clear();
   private:
    Ui::formClass ui;
    CollectXLSX cl;
    QList<QCheckBox *> check_boxes;
    QWidget *qw;
    bool onload_clicked;
   private slots:
    void onPathClick() {
        QString dir = QFileDialog::getExistingDirectory(
            this, tr("Open Directory"), ".", QFileDialog::ShowDirsOnly | QFileDialog::DontResolveSymlinks
        );
        ui.linePath->setText(dir);
    }

    void onLoadFilesClick() {
        clear();
        cl.input_dir = ui.linePath->text().toStdString();
        cl.output_name = ui.lineOut->text().toStdString();
        cl.range_1ref = ui.lineLeftCol->text().toStdString();
        cl.range_ref = ui.lineCol->text().toStdString();
        cl.sheet_index = ui.lineSheet->text().toInt();
        cl.create_input_xlsx();
        if (qw != nullptr) {
            delete qw;
        }
        qw = new QWidget();
        qw->setFixedWidth(400);
        qw->setFixedHeight((int) cl.collected.size() * 26);
        QVBoxLayout *vlayout = new QVBoxLayout();
        for (size_t i = 0; i < cl.collected.size(); ++i) {
            QCheckBox *qcheck = new QCheckBox(QString::fromStdString(cl.input_xlsx[i]), this);
            qcheck->setChecked(true);
            check_boxes.append(qcheck);
            QObject::connect(qcheck, &QCheckBox::checkStateChanged, this, &form::onCheck);
            vlayout->addWidget(qcheck);
            //ui.scrollArea->setWidget(&ui.verticalLayout);
        }
        qw->setLayout(vlayout);
        ui.scrollArea->setWidget(qw);
    }

    void onCheck() {
        QCheckBox* check = qobject_cast<QCheckBox*>(sender());
        int ind = check_boxes.indexOf(check);
        if (ind < (int) cl.collected.size() && ind >= 0) {
            cl.collected[ind] = false;
        }
        onload_clicked = true;
    }

    void onCollect() {
        try {
            if (!onload_clicked) {
                onLoadFilesClick();
            }
            cl.collect();
            QMessageBox msgBox;
            msgBox.setIcon(QMessageBox::NoIcon);
            msgBox.setWindowFlags(Qt::Dialog | Qt::CustomizeWindowHint | Qt::WindowTitleHint | Qt::WindowCloseButtonHint);
            msgBox.setWindowTitle("Готово");
            msgBox.setText(QString("Создан файл: ") + ui.lineOut->text());
            msgBox.exec();
        } catch (exception &e) {
            cerr << e.what() << endl;
        }
    }
};
