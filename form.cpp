#include "form.h"

form::form(QWidget *parent) : QMainWindow(parent) {
    ui.setupUi(this);
    QObject::connect(ui.buttonPath, &QPushButton::clicked, this, &form::onPathClick);
    QObject::connect(ui.buttonLoadFiles, &QPushButton::clicked, this, &form::onLoadFilesClick);
    QObject::connect(ui.buttonCollect, &QPushButton::clicked, this, &form::onCollect);
    onload_clicked = false;
    setFixedSize(geometry().width(), geometry().height());
}

form::~form() {
    clear();
}

void form::clear() {
    onload_clicked = false;
    cl.clear();
    for (auto check : check_boxes) {
        delete check;
    }
    check_boxes.clear();
}
